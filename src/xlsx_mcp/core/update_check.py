"""Is a newer KitchenSink4XL published? The version check behind the
update notice.

The family rules (identical in the Word, PowerPoint and Excel editors, modulo
the package and variable names):

- AT MOST ONCE A WEEK, IN THE BACKGROUND, AFTER A TOOL CALL. Nothing runs at
  import and nothing runs at server startup. After a tool call the surface
  calls maybe_refresh_in_background(), which starts ONE daemon thread when
  the cached answer is older than seven days and none was started in this
  process in the last seven days. No tool call waits for it.
  get_server_info still asks on demand through status(), with its own
  two-second timeout, and never starts a second request while the
  background one runs.
- It asks PyPI's public JSON endpoint for the package it was installed
  from. The answer is cached in a tiny JSON file under the server's own
  state directory, never beside a user's documents.
- A timeout, a network error, or a malformed payload never raises and never
  blocks a tool call; it is RECORDED and REPORTED, with the reason and the
  age of the last answer that did arrive. The report is honest about not
  knowing rather than silent.
- The PyPI payload is untrusted input. Comparison uses packaging.version,
  is prerelease-aware (a prerelease is never offered to a stable build),
  and anything unparseable yields "unknown", never a fabricated version.
- OFF SWITCH: KS4XL_UPDATE_CHECK=off disables the whole thing (no network
  call, no cache read, no cache write); the surface then says so. The
  older KS4XL_NO_UPDATE_CHECK=1 still disables it too.
- PRIVACY: the check is one plain HTTPS GET to pypi.org. It sends nothing
  but the request itself: no document, no path, no identifier, no
  telemetry. The User-Agent names the package and the installed version.
  The payload discloses this the first time it runs.
- This module never downloads or installs a new version. The notice, and
  the update a user chooses from it, live in core/update_flow.py.
"""

from __future__ import annotations

from .. import license_policy

import base64
import hashlib
import json
import os
import secrets
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from packaging.version import InvalidVersion, Version

#: PyPI distribution name (what `pip install -U` takes).
PACKAGE = "kitchensink4xl"

#: The shared editing core the launcher pins. Update now refreshes it with
#: the launcher, so a stale index page for it cannot fail the update.
CORE_PACKAGE = "kitchensink4xl-core"

#: Product name as a human reads it, used in the one surfaced line.
PRODUCT = "KitchenSink4XL"

#: The public JSON endpoint. Read-only, unauthenticated, no payload.
PYPI_URL = f"https://pypi.org/pypi/{PACKAGE}/json"

#: The off switch. The value "off" (case-insensitive) turns everything off.
OFF_ENV = "KS4XL_UPDATE_CHECK"

#: The value that means off.
OFF_VALUE = "off"

#: The original opt-out, still honored: "1"/"true"/"yes"/"on" turns it off.
LEGACY_OFF_ENV = "KS4XL_NO_UPDATE_CHECK"

#: Back-compat alias for the legacy name.
OPT_OUT_ENV = LEGACY_OFF_ENV

#: Test/ops seam: point the cache somewhere else. Undocumented on purpose.
CACHE_DIR_ENV = "KS4XL_UPDATE_CACHE_DIR"

#: How long any answer, good or bad, is trusted before asking again. One
#: real network call per week, whatever happened last time. A week, not a
#: day: these products ship small releases often, and a user does not want
#: a notice for each of them. The report always dates its own answer, so a
#: week-old "latest version" is still an honest one.
CHECK_INTERVAL = timedelta(days=7)

#: Kept as a separate name for readability; same horizon (see above).
RETRY_INTERVAL = CHECK_INTERVAL

#: The on-demand request's socket timeout (get_server_info). Two seconds.
TIMEOUT_SECONDS = 2.0

#: The background request's socket timeout. It applies to each socket
#: operation (DNS is not covered), so it is not a 30 s total; it runs on a
#: daemon thread that no tool call waits for.
BACKGROUND_TIMEOUT_SECONDS = 30.0

#: The update notices' own switch: "off" stops the notices and keeps the
#: check; "on" turns them back on after the user turned them off.
NOTICES_ENV = "KS4XL_UPDATE_NOTICES"

#: Set to "mcpb" by the Claude Desktop bundle (manifest mcp_config.env).
#: The only signal for the bundle install route.
INSTALL_ROUTE_ENV = "KS4XL_INSTALL_ROUTE"

#: What the bundle's launch argument carries between the package name and
#: the pin: "kitchensink4xl[com]==2.0.0". core/update_flow.check_bundle
#: compares the manifest to exactly this.
BUNDLE_EXTRAS = "[com]"

#: The user_config substitutions the bundle manifest sets in mcp_config.env,
#: name -> the exact manifest value (check_bundle refuses any other entry).
BUNDLE_USER_ENV = {
    "KS4XL_VERIFY_COM": "${user_config.verify_with_excel}",
    "KS4XL_ALLOWED_ROOTS": "${user_config.allowed_root}",
}

#: The GitHub repository whose releases carry the .mcpb bundle.
GITHUB_REPO = "KitchenSink4AI/KitchenSink4XL"

#: The tool that holds the update_now and update_notices ops.
UPDATE_TOOL = "xl_app"

#: This product's logger: one stderr line per event, through the server's
#: non-blocking stderr writer.
LOGGER_NAME = "kitchensink4xl"

_CACHE_NAME = "update-check.json"
_STATE_DIR_NAME = "xlsx-mcp"

# ---------------------------------------------------------------- the copy
# The sentences a caller reads out of this module.
# Machine-readable values (state names, version numbers, ISO timestamps)
# live in the payload beside them and are never restated here.

NOTE_UPDATE_AVAILABLE = (
    "A newer version is published; the numbers are beside this note. "
    "Nothing was downloaded and nothing was installed.")
NOTE_CURRENT = (
    "This build is not behind the index. PyPI answered at the time shown "
    "in last_successful_check.")
NOTE_UNKNOWN = (
    "The newest published version is not known right now, and no version "
    "is being guessed. When there is a failure reason, it is in error.")
NOTE_NOT_REACHED = (
    "PyPI could not be reached on the last attempt; the reason is in "
    "error, and the last time an answer did arrive is in "
    "last_successful_check. The check tries again after seven days. "
    "Nothing else about the server is affected.")
NOTE_DISABLED = (
    "The update check is off by configuration: the variable that turned it "
    "off is in disabled_by, and the value that does it is \"off\". While "
    "off, the server makes no network call and reads and writes no cache.")
INSTALL_NOTE = (
    "This install is pinned, so a newer version never arrives on its own. "
    "Updating means installing the new bundle, or upgrading the package "
    "with pip install -U kitchensink4xl. The server never downloads, "
    "installs, or runs anything itself.")
PRIVACY_NOTE = (
    "This check is one plain HTTPS GET to pypi.org for the package's "
    "public JSON. It sends nothing but the request itself (no document, no "
    "path, no identifier, no telemetry), and the request carries only a "
    "User-Agent naming the package and its version. Nothing else leaves the "
    "machine unless the user chooses Update now, which downloads the new "
    "version from PyPI, or the new Claude Desktop bundle from GitHub. "
    "Set KS4XL_UPDATE_CHECK=off to turn the check off.")


# ---------------------------------------------------------------- off switch


def disabled() -> bool:
    """True when the operator has turned the update check off.

    Checked BEFORE any file or network I/O in every entry point here, so a
    machine with the check off performs neither.
    """
    # The core does not know which editions poll a public index. A policy
    # that suppresses the check says so and supplies its own reason fields.
    if license_policy.update_check_suppression():
        return True
    if (os.environ.get(OFF_ENV) or "").strip().lower() == OFF_VALUE:
        return True
    return (os.environ.get(LEGACY_OFF_ENV) or "").strip().lower() in {
        "1", "true", "yes", "on",
    }


# ------------------------------------------------------------- the version


def current_version() -> str | None:
    """The running build's version, or None when it cannot be read.

    THE DISTRIBUTION'S version, which is the string get_server_info reports
    as "version" and the string a customer installed. The report used to
    carry the shared core's version here while reporting the distribution's
    version beside it, so a Business install read one number under
    current_version and another under version and a support bundle could
    not tell which build it was looking at. One string now, from the policy
    the launcher injected, with the core constant as the fallback for a
    checkout running without one.

    Deliberately NOT importlib.metadata: an editable install's recorded
    metadata goes stale the moment the version constant is bumped, and a
    wrong "you are running X" line is worse than no line at all.
    """
    try:
        version = license_policy.active().distribution_version
        if type(version) is str and version.strip():
            return version
    except Exception:
        pass
    try:
        from .. import __version__

        return str(__version__)
    except Exception:
        return None


# --------------------------------------------------------------- the cache


def cache_path() -> Path:
    """The cache file: a couple of hundred bytes in the server's own state
    directory under %LOCALAPPDATA% (the family's convention, matching the
    Word and PowerPoint siblings). Never beside a user's workbooks."""
    override = os.environ.get(CACHE_DIR_ENV)
    if override:
        return Path(override) / _CACHE_NAME
    base = (
        os.environ.get("LOCALAPPDATA")
        or os.environ.get("XDG_STATE_HOME")
        or tempfile.gettempdir()
    )
    return Path(base) / _STATE_DIR_NAME / _CACHE_NAME


def read_cache(path: Path | None = None) -> dict | None:
    """The cached answer, or None when there is not a readable one."""
    try:
        raw = (path or cache_path()).read_text(encoding="utf-8")
        data = json.loads(raw)
    except Exception:
        return None
    return data if isinstance(data, dict) else None


def atomic_write_json(target: Path, data: dict) -> bool:
    """Write JSON beside the target under a unique temporary name, then
    os.replace it into place, so two servers writing at once never leave a
    torn file. A PermissionError (Windows: another process holds the file)
    is retried a few times over about 200 ms. Returns False on failure and
    never raises: a read-only state directory must not break the server."""
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        tmp = target.with_name(f"{target.name}.{os.getpid()}.{secrets.token_hex(4)}.tmp")
        tmp.write_text(json.dumps(data), encoding="utf-8")
    except Exception:
        return False
    for attempt in range(5):
        try:
            os.replace(tmp, target)
            return True
        except PermissionError:
            time.sleep(0.02 * (attempt + 1) * 2)
        except Exception:
            break
    try:
        tmp.unlink()
    except Exception:
        pass
    return False


def write_cache(data: dict, path: Path | None = None) -> None:
    """Replace the cache file atomically. Silent on any failure (a
    read-only state directory must not break the server)."""
    try:
        atomic_write_json(path or cache_path(), data)
    except Exception:
        pass


def _parse_stamp(value) -> datetime | None:
    try:
        stamp = datetime.fromisoformat(str(value))
    except Exception:
        return None
    if stamp.tzinfo is None:
        stamp = stamp.replace(tzinfo=timezone.utc)
    return stamp


def is_due(cache: dict | None, now: datetime | None = None) -> bool:
    """Has the cached answer aged out?

    One horizon, seven days, whether the last attempt succeeded or failed:
    that is what caps the server at one real network call a week. No cache,
    or an unreadable timestamp, means due.
    """
    if not cache:
        return True
    stamp = _parse_stamp(cache.get("last_check"))
    if stamp is None:
        return True
    horizon = CHECK_INTERVAL if cache.get("ok") else RETRY_INTERVAL
    return (now or datetime.now(timezone.utc)) - stamp >= horizon


# ------------------------------------------------------------ the question


def _latest_stable(payload: dict) -> str | None:
    """The newest NON-prerelease release in a PyPI JSON payload.

    The payload is untrusted input: every field is type-checked, every
    version string goes through packaging.version, and anything that does
    not parse is skipped rather than passed along. Prefers the full
    releases map (so a prerelease published after the last stable cannot
    win), and falls back to info.version when the map is absent. Returns
    None for anything malformed.
    """
    if not isinstance(payload, dict):
        return None
    best: Version | None = None
    releases = payload.get("releases")
    if isinstance(releases, dict):
        for name, files in releases.items():
            try:
                candidate = Version(str(name))
            except InvalidVersion:
                continue
            if candidate.is_prerelease or candidate.is_devrelease:
                continue
            if not isinstance(files, list) or not all(isinstance(f, dict) for f in files):
                continue
            if isinstance(files, list):
                if not files:
                    continue  # no artifacts: not installable
                if all(
                    isinstance(f, dict) and f.get("yanked") for f in files
                ):
                    continue  # every artifact yanked
            if best is None or candidate > best:
                best = candidate
        return str(best) if best is not None else None
    info = payload.get("info")
    if isinstance(info, dict):
        try:
            candidate = Version(str(info.get("version")))
        except InvalidVersion:
            return None
        if not (candidate.is_prerelease or candidate.is_devrelease):
            return str(candidate)
    return None


def _fetch(url: str = PYPI_URL, timeout: float = TIMEOUT_SECONDS) -> dict | None:
    """One GET, no payload beyond a standard request. timeout is the socket
    timeout: two seconds on demand, BACKGROUND_TIMEOUT_SECONDS from the
    background thread. Returns the decoded JSON or None."""
    request = urllib.request.Request(
        url,
        headers={"Accept": "application/json",
                 "User-Agent": f"{PACKAGE}/{current_version() or 'unknown'}"},
    )
    with urllib.request.urlopen(request, timeout=timeout) as resp:
        body = resp.read(1_000_000)
    data = json.loads(body.decode("utf-8"))
    return data if isinstance(data, dict) else None


def _reason(exc: BaseException) -> str:
    """A short, mechanical description of why the call did not land. Carries
    only the exception class. OS and network exception messages can carry
    local paths even though the request URL is fixed."""
    # OS/proxy/TLS errors may carry local paths or credentials in their
    # message. Keep the diagnostic category, never the untrusted message.
    return type(exc).__name__[:80]


def run_check(force: bool = False, path: Path | None = None,
              timeout: float | None = None) -> str | None:
    """Perform the check if it is due, refresh the cache, and return the
    latest stable version when one is known.

    Never raises. Never prints. Called from status() (on demand, the
    two-second timeout), from the background refresh and from update_now
    (timeout given).
    """
    if disabled():
        return None
    try:
        return _run_check(force, path, timeout)
    except Exception:
        return None


def _run_check(force: bool, path: Path | None, timeout: float | None = None) -> str | None:
    target = path or cache_path()
    cache = read_cache(target)
    if not force and not is_due(cache):
        return (cache or {}).get("latest_version")
    known = (cache or {}).get("latest_version")
    last_success = (cache or {}).get("last_success")
    disclosed = bool((cache or {}).get("disclosed"))
    stamp = datetime.now(timezone.utc).isoformat()
    error: str | None = None
    try:
        # The on-demand route calls _fetch() exactly as before, so a test
        # that patches _fetch with a no-argument stand-in still works.
        payload = _fetch() if timeout is None else _fetch(timeout=timeout)
        latest = _latest_stable(payload)
        if latest is None:
            error = "unusable response: no parseable stable release"
    except Exception as exc:
        latest = None
        error = _reason(exc)
    if latest is None:
        # The attempt is recorded (that is what caps the network at one call
        # a week), the last answer that DID arrive is kept, and the reason is
        # kept with it so the surface can say what went wrong.
        write_cache({"last_check": stamp, "latest_version": known,
                     "ok": False, "last_success": last_success,
                     "error": error, "disclosed": disclosed}, target)
        return None
    write_cache({"last_check": stamp, "latest_version": latest, "ok": True,
                 "last_success": stamp, "disclosed": disclosed}, target)
    return latest


# ------------------------------------------------------------- the surface


def status(force: bool = False) -> dict:
    """The honest version report: what is running, what is published, when
    that was last confirmed, and what went wrong when it could not be.

    THIS is the only function that can reach the network, and it does so at
    most once every seven days. It never raises, and a failed or disabled
    check produces a report saying so rather than an absent one.
    """
    running = current_version()
    # The suppressing policy owns the reason it reports, so the wire shape
    # is identical to what the single-wheel build emitted.
    suppression = license_policy.update_check_suppression()
    if suppression:
        return {"state": "disabled", "current_version": running, **suppression}
    backend = resolve_backend()
    if backend.name != "pypi":
        return _channel_status(backend, force)
    if disabled():
        return {
            "state": "disabled",
            "current_version": running,
            "disabled_by": OFF_ENV,
            "off_value": OFF_VALUE,
            "note": NOTE_DISABLED,
        }

    try:
        before = read_cache() or {}
    except Exception:
        before = {}
    first_time = not before.get("disclosed")

    # Never a second request while the background one is in flight: the
    # report then comes from the cache as it stands.
    if not refresh_running():
        try:
            run_check(force=force)
        except Exception:
            pass

    try:
        cache = read_cache() or {}
    except Exception:
        cache = {}

    out: dict = {"state": "unknown", "current_version": running}
    latest = cache.get("latest_version")
    reachable = bool(cache.get("ok"))
    out["reachable"] = reachable
    last_success = cache.get("last_success")
    out["last_successful_check"] = str(last_success) if last_success else None
    if cache.get("last_check"):
        out["last_attempt"] = str(cache.get("last_check"))
    if not reachable:
        out["error"] = str(cache.get("error") or "unknown")
        out["note"] = NOTE_NOT_REACHED

    if latest:
        out["latest_version"] = str(latest)
    try:
        newer = bool(
            latest and running and Version(str(latest)) > Version(str(running))
        )
        known = bool(latest and running)
    except Exception:
        newer, known = False, False
    if not known:
        out["state"] = "unknown"
        out.setdefault("note", NOTE_UNKNOWN)
    elif newer:
        out["state"] = "update_available"
        out["install_note"] = INSTALL_NOTE
        if reachable:
            out["note"] = NOTE_UPDATE_AVAILABLE
    else:
        out["state"] = "current"
        if reachable:
            out["note"] = NOTE_CURRENT

    if first_time:
        out["privacy"] = PRIVACY_NOTE
        out["privacy_endpoint"] = PYPI_URL
        if cache:
            write_cache({**cache, "disclosed": True})

    return out


def update_notice() -> str | None:
    """One calm line when a newer stable release exists, else None.

    Reads the cache only: no network, no blocking, safe to call from any
    tool. Returns None when the check is off, when nothing is cached, when
    the versions cannot be parsed, or when the running build is current.
    """
    backend = resolve_backend()
    if backend.name != "pypi":
        return _channel_notice()
    pair = versions()
    if pair is None:
        return None
    running, latest = pair
    # The star line rides ALONG with the update notice rather than on its
    # own surface. A user reading this has an update waiting, which means
    # the project is alive and being worked on, and that is the one moment
    # asking is fair. PRODUCT is also the repository name.
    return (
        f"{PRODUCT} {running} is running; {latest} is available. "
        f"Update: pip install -U {PACKAGE} (or download the new "
        f"installer from the releases page). "
        f"Star the project: github.com/KitchenSink4AI/{PRODUCT}"
    )


def versions() -> tuple[str, str] | None:
    """(current, latest) when the cache shows a newer stable release, else
    None. The one place the comparison is made (update_notice and the
    update notice both use it).

    latest must be a stable release: never a prerelease, a dev build or a
    local version. current is the running build and may itself be a
    prerelease or a local version, so a release candidate is offered its
    final release. Both strings are packaging-normalized; nothing else from
    the cache or from PyPI is returned. Reads the cache only."""
    try:
        if disabled():
            return None
        cache = read_cache()
        if not cache:
            return None
        latest_raw = cache.get("latest_version")
        current_raw = current_version()
        if not latest_raw or not current_raw:
            return None
        latest = Version(str(latest_raw))
        running = Version(str(current_raw))
        if latest.is_prerelease or latest.is_devrelease or latest.local:
            return None
        if latest <= running:
            return None
        return str(running), str(latest)
    except Exception:
        return None


# ------------------------------------------------- the background refresh

_REFRESH_LOCK = threading.Lock()
_REFRESH_THREAD: threading.Thread | None = None
#: time.monotonic() when this process last started a background refresh.
_REFRESH_STARTED_AT: float | None = None


def refresh_running() -> bool:
    """True while this process's background refresh is in flight."""
    thread = _REFRESH_THREAD
    return thread is not None and thread.is_alive()


def maybe_refresh_in_background() -> bool:
    """The cheap guard the surface runs after every tool call. Starts the
    weekly check on a daemon thread, and returns True, only when ALL hold:
    the public channel applies (Business keeps its own resolution); the
    check is not off (disabled(): both variables and a policy that
    suppresses it); the cache is missing or older than seven days; no
    refresh is running in this process; and none was started in this
    process in the last seven days. Never raises and never blocks."""
    global _REFRESH_THREAD, _REFRESH_STARTED_AT
    try:
        if resolve_backend().name != "pypi" or disabled():
            return False
        now = time.monotonic()
        with _REFRESH_LOCK:
            if refresh_running():
                return False
            if (_REFRESH_STARTED_AT is not None
                    and now - _REFRESH_STARTED_AT < CHECK_INTERVAL.total_seconds()):
                return False
            if not is_due(read_cache()):
                return False
            thread = threading.Thread(
                target=run_check, kwargs={"timeout": BACKGROUND_TIMEOUT_SECONDS},
                name=f"{LOGGER_NAME}-update-check", daemon=True)
            _REFRESH_STARTED_AT = now
            _REFRESH_THREAD = thread
        thread.start()
        return True
    except Exception:
        return False


# ============================================================== the edition seam
# Everything above this line is the public PyPI channel and its wire
# behaviour is frozen: same URL, same two-second cap, same seven-day
# interval, same prerelease-aware comparison, same User-Agent, same cache
# file, same copy. What changed is that it is now reachable through a
# named backend, so a licensed distribution can supply a different one.
#
# A policy declares its channel through license_policy.update_backend().
# A policy that declares nothing gets PyPIBackend, which is why the
# Community wheels are untouched by construction.


class Backend:
    """What a channel has to provide. Documentation, not machinery.

    name              "pypi" or "business"
    off_env           the environment variable that turns this channel off
    privacy_text      the disclosure sentence, emitted once per cache
    privacy_endpoint  the URL that sentence names

    check() never raises, never prints, never writes outside its own cache
    path, and never downloads anything. Its return value is the raw result
    the shared module turns into status().
    """

    name: str = ""
    off_env: str = ""
    privacy_text: str = ""
    privacy_endpoint: str = ""

    def check(self, *, timeout: float, force: bool = False) -> dict:
        raise NotImplementedError

    def first_disclosure(self) -> bool:
        """Has the privacy sentence not been emitted yet on this machine?"""
        return False

    def mark_disclosed(self) -> None:
        """Remember that it has."""


class PyPIBackend(Backend):
    """The public channel, unchanged.

    Every method here delegates to the module-level functions above rather
    than re-implementing them, which is what keeps the wire behaviour and
    the existing test surface identical: a test that monkeypatches
    update_check._fetch still intercepts this backend's network call.
    """

    name = "pypi"
    off_env = OFF_ENV
    privacy_text = PRIVACY_NOTE
    privacy_endpoint = PYPI_URL

    def check(self, *, timeout: float = TIMEOUT_SECONDS,
              force: bool = False) -> dict:
        """Run the check if it is due and hand back the cache around it."""
        try:
            before = read_cache() or {}
        except Exception:
            before = {}
        run_check(force=force)
        try:
            after = read_cache() or {}
        except Exception:
            after = {}
        return {"first_time": not before.get("disclosed"), "cache": after}


# ------------------------------------------------------- the business channel

#: Which key this product occupies in a channel document's
#: product_versions map. The four products share this module and differ
#: here and in the copy constants above.
PRODUCT_KEY = "xl"

#: Where the channel document is served. An administrator may point the
#: client at an internal mirror through the policy file; the signature
#: check does not change, so a mirror cannot forge a release.
BUSINESS_ORIGIN = "https://get.kitchensink4.ai"

#: Hard cap on the business request. Three seconds or nothing.
BUSINESS_TIMEOUT_SECONDS = 3.0

#: The document is refused unread past this many bytes.
BUSINESS_MAX_BYTES = 16384

#: Default cadence, and the ceiling on the per-machine jitter that spreads
#: a fleet out so five hundred machines do not all ask at midnight.
BUSINESS_CHECK_INTERVAL_HOURS = 24
BUSINESS_MAX_JITTER_SECONDS = 7200

#: The freshness tolerance for a client clock that is wrong. Two days.
BUSINESS_CLOCK_SKEW = timedelta(hours=48)

#: THE RENEWAL SWITCH, and the only place in this client that branches on
#: renewal semantics.
#:
#: "updates_only": a release is covered when the date part of its
#: published_at is on or before the receipt's updates_through. Lapsed
#: coverage never stops the software; it only means a newer build is not
#: included.
#:
#: "term": covered when published_at is within updates_through AND the
#: client clock is before term_ends. A passed term reports
#: update_available_not_covered with reason "term_ended". The runtime
#: still disables nothing.
COVERAGE_MODEL = "updates_only"

#: The channel names, the severities, and the policy modes. Closed sets: a
#: document or a policy file carrying anything else is refused rather than
#: accommodated.
BUSINESS_CHANNELS = ("stable", "beta")
BUSINESS_SEVERITIES = ("routine", "recommended", "security")
BUSINESS_LINES = ("cp312", "cp313", "cp314")
BUSINESS_MODES = ("off", "notify", "check-on-verify")
BUSINESS_NOTIFY_TARGETS = ("console", "event-log")
DEFAULT_CHANNEL = "stable"

#: The bundle's own identity, compiled into every launch shim by the
#: installer. No product letter in the prefix: this is the bundle's
#: identity, not one product's setting.
BUNDLE_VERSION_ENV = "KS4_BUNDLE_VERSION"
RELEASE_ID_ENV = "KS4_RELEASE_ID"
CHANNEL_ENV = "KS4_CHANNEL"

#: Test and ops seams, undocumented on purpose, matching CACHE_DIR_ENV.
CONFIG_DIR_ENV = "KS4_CONFIG_DIR"
STAMP_FILE_ENV = "KS4_STAMP_FILE"

_BUSINESS_CACHE_NAME = "update-check-business.json"

#: The DER SubjectPublicKeyInfo prefix an Ed25519 public key carries when
#: an anchor file stores keys in that form rather than as raw bytes.
_ED25519_SPKI_PREFIX = bytes.fromhex("302a300506032b6570032100")

# ---------------------------------------------------------------- the copy
# The business channel carries no customer-facing sentences yet; this module
# wires the fields and names where the strings will come from. Until a
# sentence exists its field is ABSENT from the report rather than filled
# with a marker, because the marker shipped and a client read it.

#: None, not a marker string. These three carried a marker naming the copy
#: section they were waiting on, and a real MCP client read one out of a real
#: install: the note and the install note arrived inside the update_check
#: block of get_server_info, where they read as a defect in our output, which
#: is how a customer would read them too. A field with no sentence to carry
#: is omitted; a field carrying a marker says something false. When the copy
#: exists, putting a string here puts the field back with no other change.
BUSINESS_NOTE = None
BUSINESS_INSTALL_NOTE = None
BUSINESS_PRIVACY_NOTE = None


def business_copy() -> dict:
    """The copy fields a Business report carries, splatted into the report.

    Empty while the sentences are unwritten, which is what keeps the marker
    out of a customer's client.
    """
    return {} if BUSINESS_NOTE is None else {"note": BUSINESS_NOTE}


def set_install_note(report: dict) -> dict:
    """Add the install note to a report that warrants one, if there is one."""
    if BUSINESS_INSTALL_NOTE is not None:
        report["install_note"] = BUSINESS_INSTALL_NOTE
    return report


BUSINESS_UPDATE_LINE = (
    "{product} {version} is running on bundle {installed_bundle}; "
    "bundle {available_bundle} ({release_id}) is published. "
    "Your administrator installs it.")


class PolicyFileError(Exception):
    """The machine-wide update policy file exists and is not usable.

    A visible failure rather than a silent fallback to the defaults: an
    administrator who deployed a broken control needs to be told, and the
    report names the path.
    """


_REQUIRED = object()


# ------------------------------------------------------------- local inputs


def config_dir() -> Path:
    """The machine-wide KitchenSink4AI configuration root.

    Machine scope only. There is no per-user override and there will not be
    one: an update policy is a fleet control, and a file in a user-writable
    location would let the controlled user edit the control.
    """
    override = os.environ.get(CONFIG_DIR_ENV)
    if override:
        return Path(override)
    base = (os.environ.get("ProgramData") or os.environ.get("PROGRAMDATA")
            or tempfile.gettempdir())
    return Path(base) / "KitchenSink4AI"


def policy_file_path() -> Path:
    return config_dir() / "update-policy.json"


def anchor_file_path() -> Path:
    return config_dir() / "trust" / "anchors.json"


def _safe_channel(name) -> str:
    return name if name in BUSINESS_CHANNELS else DEFAULT_CHANNEL


def channel_state_path(channel: str) -> Path:
    """The anti-rollback record, one per channel, machine-wide.

    Machine-wide rather than per-server, because the highest sequence a
    machine has seen is a security fact about the machine and not about one
    of the four servers installed on it.
    """
    return config_dir() / "state" / f"channel-{_safe_channel(channel)}.json"


def business_cache_path() -> Path:
    """Beside the public channel's cache, in the server's own state
    directory. This one is a rate limiter and a report cache rather than a
    security control, so it lives with the server while the anti-rollback
    record lives with the machine."""
    return cache_path().with_name(_BUSINESS_CACHE_NAME)


def stamp_path() -> Path:
    """The environment stamp the installer writes at staging time, inside
    the versioned environment like everything else a running server
    reads."""
    override = os.environ.get(STAMP_FILE_ENV)
    if override:
        return Path(override)
    return Path(sys.prefix) / "ks4-stamp.json"


def read_stamp(path: Path | None = None) -> dict:
    """The bundle identity, environment first and file second.

    Returns bundle_version, release_id and channel, any of which may be
    None. A missing or unreadable stamp is not an error here; the caller
    reports no_stamp when the identity does not resolve.
    """
    out = {
        "bundle_version": (os.environ.get(BUNDLE_VERSION_ENV) or "").strip() or None,
        "release_id": (os.environ.get(RELEASE_ID_ENV) or "").strip() or None,
        "channel": (os.environ.get(CHANNEL_ENV) or "").strip() or None,
    }
    if all(out.values()):
        return out
    try:
        data = json.loads((path or stamp_path()).read_text(encoding="utf-8"))
    except Exception:
        return out
    if not isinstance(data, dict):
        return out
    for key in ("bundle_version", "release_id", "channel"):
        if out[key] is None:
            value = data.get(key)
            if type(value) is str and value.strip():
                out[key] = value.strip()
    return out


def _valid_release_id(value) -> bool:
    """The rule publish_release.py already enforces."""
    if type(value) is not str or not 1 <= len(value) <= 64:
        return False
    if not (value[0].isascii() and value[0].isalnum()):
        return False
    return all(c.isascii() and (c.isalnum() or c in "._-") for c in value)


def _field(data: dict, key: str, kind, default=None, allowed=None):
    """One policy field, type-checked exactly, with its documented default."""
    if key not in data or data[key] is None:
        if default is _REQUIRED:
            raise PolicyFileError(f"{key} is missing")
        return default
    value = data[key]
    if kind is int and (type(value) is not int):
        raise PolicyFileError(f"{key} is not an integer")
    if kind is str and type(value) is not str:
        raise PolicyFileError(f"{key} is not a string")
    if kind is bool and type(value) is not bool:
        raise PolicyFileError(f"{key} is not a boolean")
    if allowed is not None and value not in allowed:
        raise PolicyFileError(f"{key} is not one of {list(allowed)}")
    return value


#: The loopback hosts a policy endpoint override may name over plain http.
#: This is ks4-install.ps1's Test-UpdateEndpoint rule, and the two have to
#: agree because -UpdatePolicy validates the file and then writes it: a rule
#: the installer accepts and the client calls malformed is a policy that
#: installs cleanly and turns every server's check off, which is what the
#: gate-10 proof caught. Loopback http leaves no network, so it is not a
#: transport downgrade, and the signature check is identical either way, so a
#: mirror cannot forge a document whatever scheme it is served over.
LOOPBACK_HOSTS = ("127.0.0.1", "localhost", "[::1]")


def _valid_endpoint(origin: str) -> bool:
    """An https origin, or a loopback http origin with an optional port."""
    if origin.startswith("https://") and len(origin) > len("https://"):
        return True
    if not origin.startswith("http://"):
        return False
    rest = origin[len("http://"):].rstrip("/")
    if "/" in rest:
        return False
    if rest in LOOPBACK_HOSTS:
        return True
    # A port, if there is one. The rpartition has to run on the tail after any
    # bracketed IPv6 literal, or "[::1]" splits inside its own brackets.
    cut = rest.rfind("]") + 1 if rest.startswith("[") else 0
    head, sep, port = rest[cut:].rpartition(":")
    if not sep or not port.isdigit():
        return False
    return rest[:cut] + head in LOOPBACK_HOSTS


def read_policy_file(path: Path | None = None) -> dict:
    """The machine-wide policy, validated, with every documented default
    filled in. A missing file is normal. A malformed one raises."""
    target = path or policy_file_path()
    defaults = {
        "schema": 1,
        "mode": "notify",
        "channel": None,
        "pin": None,
        "check_interval_hours": BUSINESS_CHECK_INTERVAL_HOURS,
        "endpoint": None,
        "proxy": None,
        "notify_targets": ["console"],
        "security_bypasses_pin": False,
        "path": str(target),
        "present": False,
    }
    try:
        raw = target.read_bytes()
    except FileNotFoundError:
        return defaults
    except OSError as problem:
        raise PolicyFileError("the policy file could not be read") from problem
    try:
        # utf-8-sig, not utf-8: Notepad, Set-Content and Out-File all write
        # a BOM, the installer's ConvertFrom-Json accepts one, and a policy
        # the installer reads and the servers refuse is the worst of both.
        data = json.loads(raw.decode("utf-8-sig"))
    except Exception as problem:
        raise PolicyFileError("the policy file is not valid JSON") from problem
    if not isinstance(data, dict):
        raise PolicyFileError("the policy file is not a JSON object")
    schema = _field(data, "schema", int, 1)
    if schema != 1:
        raise PolicyFileError("schema is not 1")
    out = dict(defaults)
    out["present"] = True
    out["schema"] = 1
    out["mode"] = _field(data, "mode", str, "notify", BUSINESS_MODES)
    out["channel"] = _field(data, "channel", str, None, BUSINESS_CHANNELS)
    pin = _field(data, "pin", str, None)
    if pin is not None and not _valid_release_id(pin):
        raise PolicyFileError("pin is not a release id")
    out["pin"] = pin
    hours = _field(data, "check_interval_hours", int,
                   BUSINESS_CHECK_INTERVAL_HOURS)
    if type(hours) is bool or not 1 <= hours <= 720:
        raise PolicyFileError("check_interval_hours is outside 1 to 720")
    out["check_interval_hours"] = hours
    endpoint = _field(data, "endpoint", str, None)
    if endpoint is not None and not _valid_endpoint(endpoint):
        raise PolicyFileError(
            "endpoint is not an https origin or a loopback http origin")
    out["endpoint"] = endpoint.rstrip("/") if endpoint else None
    proxy = _field(data, "proxy", str, None)
    if proxy is not None and not proxy.startswith(("http://", "https://")):
        raise PolicyFileError("proxy is not an http or https URL")
    out["proxy"] = proxy
    targets = data.get("notify_targets")
    if targets is not None:
        if (not isinstance(targets, list)
                or not all(t in BUSINESS_NOTIFY_TARGETS for t in targets)):
            raise PolicyFileError("notify_targets carries an unknown target")
        out["notify_targets"] = list(targets)
    out["security_bypasses_pin"] = _field(
        data, "security_bypasses_pin", bool, False)
    return out


# ---------------------------------------------------------- the anchor file


def _public_key_bytes(value) -> bytes | None:
    """One anchor value normalized to the raw 32 Ed25519 key bytes."""
    if not isinstance(value, str) or not value.strip():
        return None
    text = value.strip()
    for prefix in ("ed25519:", "base64:"):
        if text.lower().startswith(prefix):
            text = text[len(prefix):]
    try:
        raw = base64.b64decode(text, validate=True)
    except Exception:
        return None
    if len(raw) == 32:
        return raw
    if len(raw) == 44 and raw.startswith(_ED25519_SPKI_PREFIX):
        return raw[-32:]
    return None


def read_anchors(path: Path | None = None) -> dict:
    """key id to raw 32 public key bytes, from the fleet anchor file.

    The file is produced by the installer and delivered out of band, so
    this reader is tolerant about its container shape and strict about the
    keys themselves: a value that is not an Ed25519 public key is dropped
    rather than half-trusted.

    utf-8-sig, like read_policy_file and for the same reason, only worse
    here. This file reaches a machine out of band and an administrator opens
    it in Notepad or writes it with Set-Content or Out-File, all of which
    leave a BOM. Decoded as plain utf-8 the BOM made json.loads raise, this
    reader returned NO ANCHORS, and every check then failed no_anchor with
    nothing anywhere saying why. A silent empty anchor set is the worst
    failure this module has: it is indistinguishable from a machine that was
    never given a trust anchor at all.
    """
    try:
        data = json.loads(
            (path or anchor_file_path()).read_text(encoding="utf-8-sig"))
    except Exception:
        return {}
    containers = []
    if isinstance(data, dict):
        for key in ("anchors", "keys", "root_keys", "public_keys"):
            if key in data:
                containers.append(data[key])
        containers.append(data)
    elif isinstance(data, list):
        containers.append(data)
    for container in containers:
        out: dict[str, bytes] = {}
        if isinstance(container, dict):
            for key_id, value in container.items():
                if not isinstance(key_id, str):
                    continue
                raw = _public_key_bytes(value)
                if raw is not None:
                    out.setdefault(key_id, raw)
        elif isinstance(container, list):
            for item in container:
                if not isinstance(item, dict):
                    continue
                key_id = item.get("key_id") or item.get("id")
                raw = _public_key_bytes(
                    item.get("public_key") or item.get("key")
                    or item.get("value"))
                if isinstance(key_id, str) and raw is not None:
                    out.setdefault(key_id, raw)
        if out:
            return out
    return {}


def _verify_ed25519(public_key: bytes, signature: bytes, message: bytes) -> bool:
    """True only when the signature verifies. False on anything else,
    including a build with no Ed25519 implementation available: an
    unverifiable document is refused, never half-trusted."""
    try:
        from cryptography.hazmat.primitives.asymmetric.ed25519 import (
            Ed25519PublicKey,
        )
    except Exception:
        return False
    try:
        Ed25519PublicKey.from_public_bytes(public_key).verify(signature, message)
        return True
    except Exception:
        return False


def canonical_bytes(payload: dict) -> bytes:
    """The signed bytes: the document with its signature member removed,
    serialized by the one rule the bundle manifest and the license receipt
    already use."""
    body = {k: v for k, v in payload.items() if k != "signature"}
    return json.dumps(body, sort_keys=True, separators=(",", ":"),
                      ensure_ascii=True).encode("utf-8")


# ------------------------------------------------------ the channel document


def _hex64(value) -> bool:
    return (type(value) is str and len(value) == 64
            and all(c in "0123456789abcdef" for c in value))


def _rfc3339(value) -> datetime | None:
    """A UTC timestamp with the Z suffix, or None."""
    if type(value) is not str or not value.endswith("Z"):
        return None
    try:
        stamp = datetime.fromisoformat(value[:-1])
    except Exception:
        return None
    return stamp.replace(tzinfo=timezone.utc)


def recalled_release_ids(document) -> list:
    """The release ids the channel head marks recalled.

    Recall status lives on the FRESH head and nowhere else: a historical
    `/release/<id>` document is byte-immutable and was signed before anyone
    knew the release would be pulled, so it can never carry its own recall.
    A head that names no list is a head that recalls nothing, which is what
    every document published before this field existed means.
    """
    if not isinstance(document, dict):
        return []
    ids = document.get("withdrawn_release_ids")
    if not isinstance(ids, list):
        return []
    return [i for i in ids if type(i) is str]


def _shape_ok(doc) -> bool:
    """Every required field present and of the documented type. Returns
    False for anything malformed; the schema number is checked by the
    caller so it can report schema_unknown separately."""
    if not isinstance(doc, dict):
        return False
    if type(doc.get("channel")) is not str:
        return False
    sequence = doc.get("sequence")
    if type(sequence) is not int or sequence < 1:
        return False
    if not _valid_release_id(doc.get("release_id")):
        return False
    for key in ("bundle_version", "default_line", "anchor_key_id",
                "anchor_fingerprint", "notes_url",
                "min_supported_bundle_version"):
        if type(doc.get(key)) is not str or not doc[key]:
            return False
    versions = doc.get("product_versions")
    if not isinstance(versions, dict) or set(versions) != {"word", "xl", "ppt", "web"}:
        return False
    if not all(type(v) is str and v for v in versions.values()):
        return False
    # Additive and optional: absent means an empty list, so a document
    # published before the field existed still validates. Present and of the
    # wrong shape is malformed rather than ignored, because a recall list
    # that is silently dropped is the one failure mode worth refusing over.
    recalled = doc.get("withdrawn_release_ids")
    if recalled is not None:
        if not isinstance(recalled, list):
            return False
        if not all(_valid_release_id(i) for i in recalled):
            return False
    lines = doc.get("lines")
    if not isinstance(lines, dict) or not lines:
        return False
    for name, entry in lines.items():
        if name not in BUSINESS_LINES or not isinstance(entry, dict):
            return False
        if not _hex64(entry.get("sha256")) or not _hex64(entry.get("manifest_sha256")):
            return False
        size = entry.get("size")
        if type(size) is not int or size < 0:
            return False
    if doc["default_line"] not in lines:
        return False
    if type(doc.get("withdrawn")) is not bool:
        return False
    reason = doc.get("withdrawn_reason")
    if doc["withdrawn"]:
        if type(reason) is not str or not reason.strip() or len(reason) > 200:
            return False
    elif reason is not None:
        return False
    if doc.get("severity") not in BUSINESS_SEVERITIES:
        return False
    if not doc["notes_url"].startswith("https://kitchensink4.ai/"):
        return False
    if _rfc3339(doc.get("published_at")) is None:
        return False
    if _rfc3339(doc.get("expires_at")) is None:
        return False
    fingerprint = doc["anchor_fingerprint"]
    if not fingerprint.startswith("sha256:") or not _hex64(fingerprint[7:]):
        return False
    signature = doc.get("signature")
    if not isinstance(signature, dict):
        return False
    if signature.get("alg") != "ed25519":
        return False
    if signature.get("key_id") != doc["anchor_key_id"]:
        return False
    if type(signature.get("sig")) is not str or not signature["sig"]:
        return False
    return True


def validate_document(raw: bytes, *, channel: str, anchors: dict,
                      highest_sequence: int = 0, now: datetime | None = None,
                      enforce_sequence: bool = True,
                      expect_release_id: str | None = None,
                      ) -> tuple[dict | None, str | None]:
    """The ordered rules of the client's anti-rollback and freshness
    contract, plus the identity check a pinned request adds. Returns
    (document, None) on acceptance, else (None, reason).

    The order matters: nothing in the document
    is used for anything before its signature verifies, and a refusal at
    any step leaves the state file untouched and never falls back to an
    unverified document.
    """
    if not isinstance(raw, (bytes, bytearray)) or len(raw) > BUSINESS_MAX_BYTES:
        return None, "oversize"
    try:
        doc = json.loads(bytes(raw).decode("utf-8"))
    except Exception:
        return None, "malformed"
    if not isinstance(doc, dict):
        return None, "malformed"
    if doc.get("schema") != 1 or type(doc.get("schema")) is not int:
        return None, "schema_unknown"
    if not _shape_ok(doc):
        return None, "malformed"
    if not anchors:
        return None, "no_anchor"
    key = anchors.get(doc["anchor_key_id"])
    if key is None:
        return None, "signature_invalid"
    try:
        signature = base64.b64decode(doc["signature"]["sig"], validate=True)
    except Exception:
        return None, "signature_invalid"
    if len(signature) != 64:
        return None, "signature_invalid"
    if not _verify_ed25519(key, signature, canonical_bytes(doc)):
        return None, "signature_invalid"
    if doc["channel"] != channel:
        return None, "channel_mismatch"
    # A signature says a document is genuine. It does not say the document
    # answers the question that was asked. A pinned request names one
    # release id, and the pin path is the one path that waives the sequence
    # floor, so an older but still validly signed document would otherwise
    # pass here in place of the pin. Identity is checked separately from
    # authenticity, and both have to hold.
    if expect_release_id is not None and doc["release_id"] != expect_release_id:
        return None, "pin_mismatch"
    moment = now or datetime.now(timezone.utc)
    if _rfc3339(doc["expires_at"]) + BUSINESS_CLOCK_SKEW < moment:
        return None, "expired"
    if enforce_sequence and doc["sequence"] < int(highest_sequence or 0):
        return None, "rollback"
    return doc, None


def read_channel_state(channel: str, path: Path | None = None) -> dict:
    """The highest sequence this machine has accepted on this channel.

    A missing or unreadable file is treated as sequence 0, not as an error,
    so a fresh machine is not broken by its absence.
    """
    try:
        data = json.loads(
            (path or channel_state_path(channel)).read_text(encoding="utf-8"))
    except Exception:
        return {"schema": 1, "highest_sequence": 0, "last_release_id": None,
                "last_seen": None}
    if not isinstance(data, dict):
        return {"schema": 1, "highest_sequence": 0, "last_release_id": None,
                "last_seen": None}
    sequence = data.get("highest_sequence")
    return {
        "schema": 1,
        "highest_sequence": sequence if type(sequence) is int and sequence >= 0 else 0,
        "last_release_id": data.get("last_release_id"),
        "last_seen": data.get("last_seen"),
    }


def write_channel_state(channel: str, data: dict, path: Path | None = None) -> None:
    """Replace the anti-rollback record atomically. Silent on failure: a
    read-only ProgramData must not break the server, and the worst case is
    that the machine forgets a sequence it already refused past."""
    target = path or channel_state_path(channel)
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        tmp = target.with_suffix(".tmp")
        tmp.write_text(json.dumps(data), encoding="utf-8")
        os.replace(tmp, target)
    except Exception:
        pass


# ----------------------------------------------------------- the local cache


def read_business_cache(path: Path | None = None) -> dict | None:
    """The business channel's own cache. Separate from the public channel's
    so a test that mines one does not trip on the other."""
    try:
        data = json.loads((path or business_cache_path()).read_text(encoding="utf-8"))
    except Exception:
        return None
    return data if isinstance(data, dict) else None


def write_business_cache(data: dict, path: Path | None = None) -> None:
    target = path or business_cache_path()
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        tmp = target.with_suffix(".tmp")
        tmp.write_text(json.dumps(data), encoding="utf-8")
        os.replace(tmp, target)
    except Exception:
        pass


# ------------------------------------------------- what the cache is good for
#
# A cached report is evidence about ONE endpoint, verified under ONE anchor
# set, for ONE install. The cache is per-user and per-product and outlives all
# three: a rebuild, an anchor rotation, a policy that moves the endpoint, or
# simply the next install on the same machine inherits the previous answer.
# The freshness rules deliberately do not re-verify a cached document's
# signature, on the reasoning that it was verified when it arrived. That holds
# only while the anchor set that admitted it is still the anchor set in force.
#
# It once stopped holding. Three servers reported an
# update_available naming a bundle version and a release id, out of a document
# signed by a scratch key generated for a proof run and fetched from a
# loopback address that had not existed for hours, on a machine whose own
# trust anchor set was EMPTY. The document was authentic once. It was not
# authentic for that machine.
#
# So every cache write records what the answer was good for, and a read whose
# binding does not match the binding in force is discarded rather than
# reported. Discarded, not refused: the check simply goes to the network, the
# way it would on a machine with no cache at all.

CACHE_BINDING_KEY = "binding"


def anchor_fingerprint(anchors: dict | None) -> str:
    """A digest of the anchor set in force, or "none" when there is none.

    Over the key ids and the raw key bytes, so a rotation that keeps a key id
    and changes the key is a different fingerprint, which is the case that
    matters: the same id verifying different bytes is precisely the state a
    cached document must not survive.
    """
    if not anchors:
        return "none"
    digest = hashlib.sha256()
    for key_id in sorted(anchors):
        value = anchors[key_id]
        digest.update(key_id.encode("utf-8"))
        digest.update(b"\x00")
        digest.update(value if isinstance(value, bytes)
                      else str(value).encode("utf-8"))
        digest.update(b"\x00")
    return digest.hexdigest()


def install_fingerprint(path: Path | None = None) -> str:
    """The install this server runs from, not the release it reports.

    The stamp's installed_at moves on every install, so two installs of the
    same release into the same environment are two identities, which is the
    pair the cache has to tell apart. The path is part of it because a second
    environment on one machine is a second install by definition.
    """
    target = path or stamp_path()
    parts = [str(target)]
    try:
        data = json.loads(target.read_text(encoding="utf-8-sig"))
    except Exception:
        data = {}
    if not isinstance(data, dict):
        data = {}
    for key in ("installed_at", "bundle_version", "release_id", "org_id"):
        value = data.get(key)
        parts.append(value.strip() if type(value) is str else "")
    return hashlib.sha256("\x00".join(parts).encode("utf-8")).hexdigest()


def cache_binding(endpoint: str | None, anchors: dict | None) -> dict:
    """The three facts a cached answer is only an answer for."""
    return {
        "endpoint": (endpoint or "").rstrip("/") or None,
        "anchors": anchor_fingerprint(anchors),
        "install": install_fingerprint(),
    }


def binding_matches(cache: dict | None, binding: dict) -> bool:
    """False for a cache with no binding at all, which is every cache written
    before this existed: an unlabelled answer cannot be shown to belong
    here."""
    if not isinstance(cache, dict):
        return False
    stored = cache.get(CACHE_BINDING_KEY)
    if not isinstance(stored, dict):
        return False
    return all(stored.get(key) == value for key, value in binding.items())


def bound_business_cache(binding: dict, path: Path | None = None) -> dict:
    """The cache, or an empty one when it was written for something else."""
    cache = read_business_cache(path)
    if cache and binding_matches(cache, binding):
        return cache
    return {}


def _machine_guid() -> str:
    """A stable per-machine string, used only to derive the local jitter.

    It is hashed into a number of seconds and stored in the cache file. It
    is never sent anywhere, never reported, and never reaches the request:
    the whole point of the jitter is that a fleet spreads itself out
    WITHOUT the server learning who asked.
    """
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r"SOFTWARE\Microsoft\Cryptography") as handle:
            value, _ = winreg.QueryValueEx(handle, "MachineGuid")
            if isinstance(value, str) and value.strip():
                return value.strip()
    except Exception:
        pass
    try:
        # socket.gethostname() is the value platform.node() returns, without
        # platform.uname(), which on Windows asks WMI and can hang.
        import socket

        return socket.gethostname() or "unknown-machine"
    except Exception:
        return "unknown-machine"


def business_jitter(channel: str, cache: dict | None = None) -> int:
    """Zero to 7200 seconds, deterministic per machine and channel.

    Computed once and kept in the cache file, so it is stable across calls
    and needs no randomness at call time.
    """
    if isinstance(cache, dict):
        stored = cache.get("jitter_seconds")
        if type(stored) is int and 0 <= stored < BUSINESS_MAX_JITTER_SECONDS:
            return stored
    digest = hashlib.sha256(
        (_machine_guid() + channel).encode("utf-8")).digest()
    return int.from_bytes(digest[:8], "big") % BUSINESS_MAX_JITTER_SECONDS


def business_is_due(cache: dict | None, interval_hours: int, jitter: int,
                    now: datetime | None = None) -> bool:
    """Has the cached answer aged out? One horizon, whatever happened last
    time, which is what caps the network at one call per interval."""
    if not cache:
        return True
    stamp = _parse_stamp(cache.get("last_check"))
    if stamp is None:
        return True
    horizon = timedelta(hours=interval_hours) + timedelta(seconds=jitter)
    return (now or datetime.now(timezone.utc)) - stamp >= horizon


# ---------------------------------------------------------------- the call


class _SameOriginRedirects(urllib.request.HTTPRedirectHandler):
    """At most one redirect, and only to the same origin."""

    max_redirections = 1

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        from urllib.parse import urlsplit

        here, there = urlsplit(req.full_url), urlsplit(newurl)
        if (there.scheme, there.netloc) != (here.scheme, here.netloc):
            return None
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _business_open(request, timeout: float, proxy: str | None):
    """The one network call this channel makes. Split out so a test can
    replace it and read back the whole request."""
    handlers = [_SameOriginRedirects()]
    if proxy:
        handlers.append(urllib.request.ProxyHandler({"http": proxy,
                                                     "https": proxy}))
    else:
        handlers.append(urllib.request.ProxyHandler({}))
    opener = urllib.request.build_opener(*handlers)
    return opener.open(request, timeout=timeout)


def business_fetch(url: str, *, user_agent: str,
                   timeout: float = BUSINESS_TIMEOUT_SECONDS,
                   proxy: str | None = None) -> bytes:
    """One HTTPS GET. Exactly two headers, no query string, no cookie, no
    identifier of any kind, and a hard byte cap on the body.

    Raises on failure; the caller turns the exception into a reason.
    """
    request = urllib.request.Request(
        url, method="GET",
        headers={"Accept": "application/json", "User-Agent": user_agent})
    with _business_open(request, timeout, proxy) as response:
        return response.read(BUSINESS_MAX_BYTES + 1)


def _failure_reason(exc: BaseException) -> str:
    """Weather, or something an operator should act on."""
    if isinstance(exc, TimeoutError):
        return "timeout"
    reason = getattr(exc, "reason", None)
    if isinstance(reason, TimeoutError):
        return "timeout"
    if type(exc).__name__ in ("timeout", "socket.timeout"):
        return "timeout"
    return "network"


# ------------------------------------------------------------- the coverage


def _date_part(stamp: str) -> date | None:
    try:
        return date.fromisoformat(str(stamp)[:10])
    except Exception:
        return None


def evaluate_coverage(published_at: str, updates_through, term_ends,
                      now: datetime | None = None,
                      model: str | None = None) -> dict:
    """Does this customer's coverage include this release?

    The one function in the client that branches on renewal semantics.
    Returns the coverage block the report carries, with covered None when
    there is no receipt to decide from, which is "does not know" rather
    than "no".
    """
    chosen = model or COVERAGE_MODEL
    out = {"model": chosen, "covered": None,
           "updates_through": updates_through or None,
           "term_ends": term_ends or None, "reason": None}
    published = _date_part(published_at)
    through = _date_part(updates_through) if updates_through else None
    if published is None or through is None:
        return out
    covered = published <= through
    if not covered:
        out["covered"] = False
        out["reason"] = "updates_lapsed"
        return out
    if chosen == "term":
        ends = _date_part(term_ends) if term_ends else None
        if ends is not None:
            today = (now or datetime.now(timezone.utc)).date()
            if today > ends:
                out["covered"] = False
                out["reason"] = "term_ended"
                return out
    out["covered"] = True
    return out


# ------------------------------------------------------------- the backend


#: The Business channel's own off switch, machine-wide, and deliberately NOT
#: the per-product community variable.
#:
#: ks4-install.ps1 compiles `<PREFIX>_UPDATE_CHECK=off` into every Business
#: launch shim. That is what keeps a core paired with a community or a missing
#: adapter from ever reaching the public index, and it is not going away. A
#: Business backend that honoured the same name is therefore off on every real
#: install: a freshly installed bundle reported check_disabled with
#: disabled_by set to the product's own variable. So the shim keeps the
#: community switch and "the business check is governed by the policy
#: file alone". This is the one environment name that still turns the Business
#: channel off, and it is the same name ks4-install.ps1 reads for its own
#: -CheckUpdates, so one variable turns the whole channel off on a machine.
BUSINESS_OFF_ENV = "KS4_UPDATE_CHECK"


def _business_env_off() -> bool:
    """The environment can only turn the Business check OFF, never on."""
    return (os.environ.get(BUSINESS_OFF_ENV) or "").strip().lower() == OFF_VALUE


class BusinessBackend(Backend):
    """The licensed channel: one identifier-free GET for a signed list of
    published releases, verified against the anchor the installer pinned.

    It reads the policy file, the anchor file, the environment stamp and
    the verified receipt locally, sends none of them, downloads nothing,
    installs nothing, never raises, and never prints. When it cannot reach
    the endpoint it says it does not know rather than reporting "current"
    forever.
    """

    name = "business"
    off_env = BUSINESS_OFF_ENV
    privacy_text = BUSINESS_PRIVACY_NOTE
    privacy_endpoint = BUSINESS_ORIGIN + "/release/channel/"

    def __init__(self, selector=None):
        #: Whatever the policy handed back. Kept for its declared channel
        #: and for diagnostics; the check reads its own inputs.
        self.selector = selector

    def first_disclosure(self) -> bool:
        return not (read_business_cache() or {}).get("disclosed")

    def mark_disclosed(self) -> None:
        cache = read_business_cache()
        if cache:
            write_business_cache({**cache, "disclosed": True})

    # ------------------------------------------------------------ inputs

    def user_agent(self) -> str:
        """Product and version only. No org, no license, no machine."""
        try:
            policy = license_policy.active()
            return f"{policy.product_id}/{policy.distribution_version}"
        except Exception:
            return f"{PACKAGE}-business/{current_version() or 'unknown'}"

    def license_facts(self) -> dict:
        """updates_through and term_ends out of the verified receipt.

        Redacted: the licensee's name is never read here and never reaches
        the report.
        """
        try:
            described = license_policy.active().describe_license(redacted=True)
        except Exception:
            return {"updates_through": None, "term_ends": None}
        if not isinstance(described, dict):
            return {"updates_through": None, "term_ends": None}
        return {"updates_through": described.get("updates_through"),
                "term_ends": described.get("term_ends")}

    def resolve_channel(self, stamp: dict, policy: dict) -> str:
        for candidate in (stamp.get("channel"), policy.get("channel")):
            if candidate in BUSINESS_CHANNELS:
                return candidate
        return DEFAULT_CHANNEL

    def document_url(self, channel: str, policy: dict,
                     release_id: str | None = None) -> str:
        origin = policy.get("endpoint") or BUSINESS_ORIGIN
        if release_id:
            return f"{origin}/release/{release_id}"
        return f"{origin}/release/channel/{channel}"

    # -------------------------------------------------------------- check

    def check(self, *, timeout: float = BUSINESS_TIMEOUT_SECONDS,
              force: bool = False) -> dict:
        """The whole business check, start to finish. Never raises."""
        try:
            return self._check(timeout=timeout, force=force)
        except Exception:
            # A defect here is still not allowed to break a status call.
            return _business_failure("network", installed=None, source=None)

    def _check(self, *, timeout: float, force: bool) -> dict:
        stamp = read_stamp()
        installed = {"bundle_version": stamp["bundle_version"],
                     "release_id": stamp["release_id"],
                     "channel": stamp["channel"]}

        # Precedence, highest first: the environment can only turn the
        # check OFF, never on.
        if _business_env_off():
            return {"status": "check_disabled", "installed": installed,
                    "available": None, "coverage": None, "pin": None,
                    "checked_at": _now_stamp(), "last_successful_check": None,
                    "source": None, "reason": None,
                    "disabled_by": BUSINESS_OFF_ENV, "off_value": OFF_VALUE,
                    **business_copy()}
        try:
            policy = read_policy_file()
        except PolicyFileError as problem:
            failed = _business_failure("policy_malformed", installed=installed,
                                       source=None)
            failed["policy_path"] = str(policy_file_path())
            failed["policy_error"] = str(problem)
            return failed
        if policy["mode"] == "off":
            # No network call, no cache read, no cache write.
            return {"status": "check_disabled", "installed": installed,
                    "available": None, "coverage": None,
                    "pin": policy["pin"], "checked_at": _now_stamp(),
                    "last_successful_check": None, "source": None,
                    "reason": None, "disabled_by": "update-policy.json",
                    "policy_path": policy["path"], **business_copy()}

        channel = self.resolve_channel(stamp, policy)
        installed["channel"] = installed["channel"] or channel
        if not installed["release_id"] or not installed["bundle_version"]:
            # An honest refusal: this environment was not installed from a
            # release, so there is nothing to compare against.
            return _business_failure("no_stamp", installed=installed,
                                     source=None)

        # The anchor set and the endpoint are read BEFORE the cache, because
        # a cached answer cannot be judged without knowing what it would have
        # to match. A cache written under another endpoint, another anchor set
        # or another install is not this machine's answer and is dropped here,
        # which sends the check to the network exactly as a first run does.
        anchors = read_anchors()
        binding = cache_binding(policy.get("endpoint") or BUSINESS_ORIGIN,
                                anchors)
        cache = bound_business_cache(binding)
        jitter = business_jitter(channel, cache)
        due = force or business_is_due(cache, policy["check_interval_hours"],
                                       jitter)
        source = self.document_url(channel, policy)

        if not due:
            document = cache.get("document")
            cached_pin = cache.get("pin_document")
            # A refresh that failed stays failed until one succeeds. The
            # document the last good check left behind is evidence, not an
            # answer: reporting from it would turn a check that could not be
            # made into the word "current", which is the one thing this
            # module must never say when it does not know.
            if not cache.get("ok", False):
                return _business_failure(cache.get("error") or "network",
                                         installed=installed,
                                         source=cache.get("source") or source,
                                         last_success=cache.get("last_success"))
            if isinstance(document, dict):
                # A pin set or changed since the cache was written is a
                # different question from the one the cache answers.
                if not self._cache_answers_the_pin(document, cached_pin, policy):
                    return _business_failure("pin_mismatch", installed=installed,
                                             source=cache.get("source") or source,
                                             last_success=cache.get("last_success"))
                # A cached document was authenticated when it arrived and
                # it stays authentic, but authenticity is not freshness. The
                # polling interval can outlast the document's own
                # expires_at, and once it has, the cached answer is past the
                # lifetime it was signed for. The same rules the network
                # path applies are applied here, with the same clock
                # tolerance and the same sequence floor, read from disk, so
                # this costs no network call and does not change the poll
                # interval. Reporting an expired document as current or as
                # update_available is the failure this prevents.
                stale = self._stale_cache_reason(
                    document, cached_pin, channel,
                    read_channel_state(channel)["highest_sequence"])
                if stale is not None:
                    return _business_failure(
                        stale, installed=installed,
                        source=cache.get("source") or source,
                        last_success=cache.get("last_success"))
                return self._report(document, installed, policy, channel,
                                    source=cache.get("source") or source,
                                    last_success=cache.get("last_success"),
                                    pin_document=cached_pin)
            return _business_failure(cache.get("error") or "network",
                                     installed=installed, source=source,
                                     last_success=cache.get("last_success"))

        state = read_channel_state(channel)
        head, pin_document = None, None

        # The head is fetched on EVERY check, pinned or not. A pin resolves a
        # historical document, and a historical document is immutable: it was
        # signed before anyone knew the release would be pulled, so it cannot
        # carry its own recall. Only the fresh head can say that a pinned
        # release has been withdrawn, which makes the head the one document a
        # pinned fleet most needs. A head that cannot be verified means this
        # machine does not know, and not knowing is never reported as current.
        head, reason = self._resolve(source, channel, anchors,
                                     state["highest_sequence"],
                                     policy, timeout,
                                     enforce_sequence=True)
        if head is None:
            return self._record_failure(cache, reason, installed, source,
                                        jitter, channel, binding)

        # One clock reading for the whole of a successful check, taken the
        # moment the head is accepted and carried from here to every record
        # and every field that dates this observation. Reading the clock
        # again for any of them can only introduce a difference that means
        # nothing, because a check does not succeed twice a second apart,
        # and it was exactly such a difference that made the expiry test
        # report one second in the cache and the next one in its answer.
        # The two dates in a report can still disagree, and that is the
        # signal worth keeping: it says this check did not succeed and what
        # you are reading came from an older one.
        stamp_now = _now_stamp()
        if head["sequence"] > state["highest_sequence"]:
            write_channel_state(channel, {
                "schema": 1, "highest_sequence": head["sequence"],
                "last_release_id": head["release_id"],
                "last_seen": stamp_now})

        if policy["pin"] is not None:
            bypass = (policy["security_bypasses_pin"]
                      and head.get("severity") == "security")
            if not bypass:
                pin_url = self.document_url(channel, policy,
                                            release_id=policy["pin"])
                # A pin names a release that may legitimately be OLDER than
                # the channel head, so the sequence rule is not applied to
                # it and it never advances the anti-rollback record. Every
                # other rule, signature included, still is.
                pin_document, reason = self._resolve(
                    pin_url, channel, anchors, 0, policy, timeout,
                    enforce_sequence=False,
                    expect_release_id=policy["pin"])
                if pin_document is None:
                    return self._record_failure(cache, reason, installed,
                                                pin_url, jitter, channel,
                                                binding)
                source = pin_url

        write_business_cache({
            "last_check": stamp_now, "ok": True, "error": None,
            "jitter_seconds": jitter, "last_success": stamp_now,
            "source": source, "document": head,
            "pin_document": pin_document,
            "disclosed": bool(cache.get("disclosed")),
            CACHE_BINDING_KEY: binding,
        })
        return self._report(head, installed, policy, channel, source=source,
                            last_success=stamp_now,
                            pin_document=pin_document,
                            checked_at=stamp_now)

    # ------------------------------------------------------------ helpers

    def _resolve(self, url, channel, anchors, highest, policy, timeout,
                 *, enforce_sequence, expect_release_id=None):
        try:
            raw = business_fetch(url, user_agent=self.user_agent(),
                                 timeout=timeout, proxy=policy.get("proxy"))
        except Exception as problem:
            return None, _failure_reason(problem)
        return validate_document(raw, channel=channel, anchors=anchors,
                                 highest_sequence=highest,
                                 enforce_sequence=enforce_sequence,
                                 expect_release_id=expect_release_id)

    @staticmethod
    def _cache_answers_the_pin(document, pin_document, policy) -> bool:
        """Whether a cached pair still answers the policy in force now.

        A pin that was set, cleared or repointed after the cache was written
        makes the cached answer an answer to a different question."""
        pin = policy.get("pin")
        if pin is None:
            return pin_document is None
        bypass = (policy.get("security_bypasses_pin")
                  and isinstance(document, dict)
                  and document.get("severity") == "security")
        if bypass:
            return True
        return (isinstance(pin_document, dict)
                and pin_document.get("release_id") == pin)

    @staticmethod
    def _stale_cache_reason(document, pin_document, channel, highest_sequence,
                            now=None):
        """The reason a cached pair can no longer be reported, or None.

        A signature is checked once, when the bytes arrive, and it never
        stops being valid. Freshness does stop: every document carries its
        own expires_at, and the cache can be read long after that moment
        has passed. These are the network path's rules 4 through 6,
        re-applied locally to bytes already on disk, in the
        same order and with the same 48 hour clock tolerance. The signature
        is not re-verified because it was verified before the document was
        ever written to the cache.
        """
        moment = now or datetime.now(timezone.utc)
        # The head carries the sequence floor; a pinned document is allowed
        # to sit below it, which is the one rule a pin waives, and is the
        # same exemption the network path gives it.
        for doc, enforce_sequence in ((document, True), (pin_document, False)):
            if not isinstance(doc, dict):
                continue
            if doc.get("channel") != channel:
                return "channel_mismatch"
            expires = _rfc3339(doc.get("expires_at"))
            if expires is None or expires + BUSINESS_CLOCK_SKEW < moment:
                return "expired"
            sequence = doc.get("sequence")
            if not isinstance(sequence, int) or isinstance(sequence, bool):
                return "malformed"
            if enforce_sequence and sequence < int(highest_sequence or 0):
                return "rollback"
        return None

    def _record_failure(self, cache, reason, installed, source, jitter,
                        channel, binding=None):
        """Record the attempt so the interval still caps the network, keep
        whatever document did arrive last, and leave the anti-rollback
        record untouched.

        The binding is the one in force now. A failed check writes it for the
        same reason a good one does: the record it leaves behind carries a
        document, and a document is only evidence for the machine it was
        fetched for.
        """
        record = {
            "last_check": _now_stamp(), "ok": False, "error": reason,
            "jitter_seconds": jitter,
            "last_success": cache.get("last_success"),
            "source": cache.get("source") or source,
            "document": cache.get("document"),
            "pin_document": cache.get("pin_document"),
            "disclosed": bool(cache.get("disclosed")),
        }
        if binding is not None:
            record[CACHE_BINDING_KEY] = binding
        write_business_cache(record)
        return _business_failure(reason, installed=installed, source=source,
                                 last_success=cache.get("last_success"))

    def _report(self, document, installed, policy, channel, *, source,
                last_success, pin_document=None, checked_at=None):
        """Turn an accepted document into one of the seven statuses.

        ``checked_at`` is when THIS check observed the channel, which is not
        always when the last good one did: a cached or failed answer reports
        the moment it was asked alongside the older moment it last knew
        anything. A caller that has already captured the instant of a
        successful check hands it in here so the report and the cache carry
        one timestamp rather than two readings of a clock that ticked
        between them.
        """
        target = pin_document or document
        facts = self.license_facts()
        coverage = evaluate_coverage(target["published_at"],
                                     facts["updates_through"],
                                     facts["term_ends"])
        available = {
            "bundle_version": target["bundle_version"],
            "release_id": target["release_id"],
            "published_at": target["published_at"],
            "severity": target["severity"],
            "notes_url": target["notes_url"],
            "product_versions": dict(target["product_versions"]),
        }
        out = {
            "status": "current",
            "installed": installed,
            "available": None,
            "coverage": coverage,
            "pin": policy["pin"],
            "checked_at": checked_at or _now_stamp(),
            "last_successful_check": last_success,
            "source": source,
            "reason": None,
            **business_copy(),
        }
        # A recall is carried by the fresh head, so it reaches a pinned
        # machine that would otherwise never look at the head at all. The
        # target is the pin when there is one and the install when there is
        # not, which is the release this report is actually about.
        recalled = recalled_release_ids(document)
        out["recalled_release_ids"] = list(recalled)
        if target["release_id"] in recalled or installed["release_id"] in recalled:
            out["status"] = "withdrawn"
            out["available"] = available
            out["reason"] = "withdrawn"
            out["withdrawn_reason"] = document.get("withdrawn_reason")
            set_install_note(out)
            return out
        if document.get("withdrawn") and document["release_id"] == installed["release_id"]:
            out["status"] = "withdrawn"
            out["available"] = available
            out["reason"] = "withdrawn"
            out["withdrawn_reason"] = document.get("withdrawn_reason")
            set_install_note(out)
            return out
        if target["release_id"] == installed["release_id"]:
            return out
        if document.get("withdrawn") and pin_document is None:
            # The published head has been pulled and it is not what is
            # installed, so there is nothing to move to.
            out["withdrawn_release_id"] = document["release_id"]
            return out
        out["available"] = available
        set_install_note(out)
        if coverage["covered"] is False:
            out["status"] = "update_available_not_covered"
            out["reason"] = coverage["reason"]
        elif target["severity"] == "security":
            out["status"] = "security_update"
        else:
            out["status"] = "update_available"
        return out


def _env_off() -> bool:
    """The environment can only turn the check OFF, never on."""
    if (os.environ.get(OFF_ENV) or "").strip().lower() == OFF_VALUE:
        return True
    return (os.environ.get(LEGACY_OFF_ENV) or "").strip().lower() in {
        "1", "true", "yes", "on"}


def _now_stamp() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace(
        "+00:00", "Z")


def _business_failure(reason: str, *, installed, source,
                      last_success=None) -> dict:
    """check_failed, with the reason named. The server is unaffected, and
    no sentence saying so is shipped yet, so the
    report carries the reason and no sentence at all."""
    return {
        "status": "check_failed",
        "installed": installed,
        "available": None,
        "coverage": None,
        "pin": None,
        "checked_at": _now_stamp(),
        "last_successful_check": last_success,
        "source": source,
        "reason": reason,
        **business_copy(),
    }


# ------------------------------------------------------------- the dispatch


def resolve_backend() -> Backend:
    """Which channel this distribution uses.

    A policy that declares nothing gets the public channel, which is why
    the Community wheels are untouched by construction. A policy may hand
    back a full backend of its own, or a selector naming one this module
    implements; the second shape is what the Business adapter uses, because
    that adapter wheel is product-neutral and imports nothing from any
    editing core, so it cannot construct a backend that lives in one.
    """
    try:
        selector = license_policy.update_backend()
    except Exception:
        selector = None
    if selector is None:
        return PyPIBackend()
    if callable(getattr(selector, "check", None)):
        return selector
    if getattr(selector, "backend_id", None) == "business":
        return BusinessBackend(selector)
    return PyPIBackend()


def business_active() -> bool:
    """True when this installation runs the licensed channel."""
    return isinstance(resolve_backend(), BusinessBackend)


def installed_build() -> dict:
    """bundle_version, release_id and channel for the report, or three
    nulls on a public-channel install. A support bundle tells the two
    editions apart at a glance by exactly this."""
    if not business_active():
        return {"bundle_version": None, "release_id": None, "channel": None}
    stamp = read_stamp()
    return {"bundle_version": stamp["bundle_version"],
            "release_id": stamp["release_id"],
            "channel": stamp["channel"]}


# ------------------------------------------------- the non-public surface


def _channel_status(backend: Backend, force: bool = False) -> dict:
    """status() for any channel that is not the public index.

    The backend produces the report; this function owns the one thing every
    channel shares, which is the once-per-cache privacy disclosure. A
    disabled check discloses nothing, because a check that is off has
    nothing to disclose.
    """
    try:
        out = backend.check(timeout=BUSINESS_TIMEOUT_SECONDS, force=force)
    except Exception:
        out = _business_failure("network", installed=None, source=None)
    if not isinstance(out, dict):
        out = _business_failure("malformed", installed=None, source=None)
    out.setdefault("current_version", current_version())
    if out.get("status") == "check_disabled":
        return out
    try:
        if backend.privacy_text and backend.first_disclosure():
            out["privacy"] = backend.privacy_text
            out["privacy_endpoint"] = backend.privacy_endpoint
            backend.mark_disclosed()
    except Exception:
        pass
    return out


def _channel_notice() -> str | None:
    """One calm line when a newer bundle is published, else None.

    Reads the cache only: no network, no blocking, safe to call from any
    tool. A check that is off, a stamp that does not resolve, and a cache
    that holds nothing all produce None rather than a guess.
    """
    try:
        if _business_env_off():
            return None
        try:
            policy = read_policy_file()
        except PolicyFileError:
            return None
        if policy["mode"] == "off":
            return None
        # The same binding rule as the check. This line is the one a tool
        # prints without asking anybody, so a cached answer from another
        # endpoint, another anchor set or another install must not reach it
        # either.
        cache = bound_business_cache(cache_binding(
            policy.get("endpoint") or BUSINESS_ORIGIN, read_anchors()))
        document = cache.get("pin_document") or cache.get("document")
        if not isinstance(document, dict):
            return None
        stamp = read_stamp()
        if not stamp["release_id"] or not stamp["bundle_version"]:
            return None
        if document.get("release_id") == stamp["release_id"]:
            return None
        policy_object = license_policy.active()
        return BUSINESS_UPDATE_LINE.format(
            product=PRODUCT,
            version=policy_object.distribution_version,
            installed_bundle=stamp["bundle_version"],
            available_bundle=document.get("bundle_version"),
            release_id=document.get("release_id"))
    except Exception:
        return None
