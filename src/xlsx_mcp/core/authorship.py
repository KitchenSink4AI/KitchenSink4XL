"""Who the notes this server writes are signed by.

The user's choice, with the user's own name as the default, as in the Word
editor:

- A call's own `author` argument wins, unless it is an email address. The
  value "assistant" (any case, spaces around it ignored) is reserved: it
  signs ASSISTANT_LABEL, as the setting below does, whatever the setting is.
  Any other name is used as given.
- Otherwise the launch setting KS4XL_AUTHOR decides. `user` (the default)
  signs with the user's Office user name, read from
  HKCU\\Software\\Microsoft\\Office\\Common\\UserInfo. `assistant` signs
  ASSISTANT_LABEL.
- With no Office user name set, the Windows sign-in name (the USERNAME,
  USER or LOGNAME environment variable) stands in; without that either,
  FALLBACK_NAME.

A value that is an email address (it contains "@") is never written as an
author, whoever supplies it: the call's own author, or the Office and Windows
names. The next place in that order stands in for it, and when the call sent
the address, the answer says so in one sentence (signer_with_note).

The answer of a write says who signed and where the name came from
(author_source), so a caller can tell the user.
"""

from __future__ import annotations

import os
import sys

#: The label the assistant setting signs with.
ASSISTANT_LABEL = "AI Assistant via KS4XL"

#: The name used when neither Office nor Windows names the user.
FALLBACK_NAME = "Author"

#: KS4XL_AUTHOR's values.
USER, ASSISTANT = "user", "assistant"
MODES = (USER, ASSISTANT)

#: Where a signing name came from, reported as author_source.
EXPLICIT = "explicit"                # the call named an author
EXPLICIT_ASSISTANT = "explicit_assistant"  # the call asked for the label: author="assistant"
SETTING = "setting"                  # KS4XL_AUTHOR=assistant chose the label
OFFICE_USERNAME = "office_username"  # the Office user name
WINDOWS_SIGNIN = "windows_signin"    # the Windows sign-in name stood in
FALLBACK = "fallback"                # neither named the user
SOURCES = (EXPLICIT, EXPLICIT_ASSISTANT, SETTING, OFFICE_USERNAME, WINDOWS_SIGNIN, FALLBACK)

#: What an email address was replaced in: a note's signature, or the
#: workbook's author property. It sets the wording of the sentence.
NOTE, PROPERTY = "note", "property"

_USER_INFO = r"Software\Microsoft\Office\Common\UserInfo"


def _office_value(name: str) -> str | None:
    """A value of the Office user information, read only, or None."""
    if sys.platform != "win32":
        return None
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _USER_INFO) as key:
            value, vtype = winreg.QueryValueEx(key, name)
    except OSError:
        return None
    if vtype not in (1, 2) or not isinstance(value, str):  # REG_SZ, REG_EXPAND_SZ
        return None
    value = value.strip()
    return value or None


def asks_for_assistant(author: object) -> bool:
    """True for the reserved author value: "assistant", in any case, with
    spaces around it ignored."""
    return isinstance(author, str) and author.strip().casefold() == ASSISTANT


def _a_name(value: str | None) -> str | None:
    """``value`` when it can sign: not blank and not an email address."""
    value = (value or "").strip()
    if not value or "@" in value:
        return None
    return value


def _user_identity() -> tuple[str, str]:
    """(name, source): the user's own name, and where it was read from. An
    Office or Windows name that is an email address is passed over."""
    name = _a_name(_office_value("UserName"))
    if name:
        return name, OFFICE_USERNAME
    for var in ("USERNAME", "USER", "LOGNAME"):
        value = _a_name(os.environ.get(var))
        if value:
            return value, WINDOWS_SIGNIN
    return FALLBACK_NAME, FALLBACK


def mode() -> str:
    """The KS4XL_AUTHOR value this server started with ('user' by default)."""
    from .. import runtime

    try:
        return runtime.plan().settings.author_mode
    except Exception:  # noqa: BLE001 - no plan, no setting: the default
        return USER


#: How each source reads in the sentence that says an email address was set aside.
_SOURCE_WORDS = {
    SETTING: "the assistant label KS4XL_AUTHOR=assistant chooses",
    OFFICE_USERNAME: "the user's Office user name",
    WINDOWS_SIGNIN: "the user's Windows sign-in name",
    FALLBACK: "the name used when neither Office nor Windows names the user",
}


def _email_note(name: str, source: str, what: str) -> str:
    """The one sentence an answer carries when the call sent an email address
    as the author and it was not written. It never repeats the address."""
    done = "this note is signed" if what == NOTE else "the author is set to"
    return (f'You sent an email address as the author, which is never written into a '
            f'workbook, so {done} "{name}", {_SOURCE_WORDS.get(source, source)}, instead.')


def signer_with_note(author: str | None, what: str = NOTE) -> tuple[str, str, str | None]:
    """(name, source, note or None) for the author a call may have named.

    "assistant" asks for the assistant label. Any other author the call names
    wins unless it is an email address; then the next place in the usual order
    stands in: the assistant label with KS4XL_AUTHOR=assistant, else the
    user's own name. The note is the sentence the answer carries when the
    address was set aside."""
    if asks_for_assistant(author):
        return ASSISTANT_LABEL, EXPLICIT_ASSISTANT, None
    if author and "@" not in str(author):
        return author, EXPLICIT, None
    if mode() == ASSISTANT:
        name, source = ASSISTANT_LABEL, SETTING
    else:
        name, source = _user_identity()
    return name, source, (_email_note(name, source, what) if author else None)


def signer(author: str | None) -> tuple[str, str]:
    """(name, source) a note signs with: the call's own author, else the
    launch setting's label, else the user's own name."""
    return signer_with_note(author)[:2]


__all__ = ["ASSISTANT_LABEL", "FALLBACK_NAME", "USER", "ASSISTANT", "MODES", "SOURCES",
           "NOTE", "PROPERTY", "asks_for_assistant", "mode", "signer", "signer_with_note"]
