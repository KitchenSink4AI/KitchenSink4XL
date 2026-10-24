"""Breaking a stale lockfile has one winner.

A waiter that finds a lockfile whose holder is gone deletes it and publishes
its own. Two waiters that had both read the same dead holder's lockfile used
to both delete "the lockfile": the slower one deleted the LIVE lock the faster
one had just published, both then held the lock, and two saves that both
passed the stamp check before either promoted lost an update (measured: two
holders in 1 run of 10 with a 250 ms pause between the read and the delete).

core/lockrelease.py break_stale now breaks under an OS lock on a guard file
beside the lockfile, reads the lockfile again, and deletes it only while it is
still the very file it judged. The tests drive real processes.

Portable: the same file runs in every editor that keeps such a lockfile. Only
the import line and the three module names below name this product's modules.
"""

from __future__ import annotations

import importlib
import json
import os
import subprocess
import sys
import textwrap
import time
from pathlib import Path

import pytest

from xlsx_mcp.core import lockrelease

#: Every module that keeps an advisory lockfile and breaks a stale one.
LOCK_MODULES = ["xlsx_mcp.core.safesave"]
#: The module whose write_lock the process tests drive.
FILE_LOCK = "xlsx_mcp.core.safesave"

locks = importlib.import_module(FILE_LOCK)
# The src folder of the package this test imported, so a child process runs
# the same code as the test.
SRC = Path(lockrelease.__file__).resolve().parents[2]
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _dead_pid() -> int:
    p = subprocess.Popen([sys.executable, "-c", "pass"], stdin=subprocess.DEVNULL,
                         creationflags=NO_WINDOW)
    p.wait()
    return p.pid


def _stale_record() -> dict:
    dead = _dead_pid()
    return {"pid": dead, "token": f"{dead}-feedface", "time": time.time() - 30,
            "host": locks._local_host()}


def _guard_of(lock: Path) -> Path:
    return lock.with_name(lock.name + lockrelease.GUARD_SUFFIX)


# ------------------------------------------------- one breaker, in process


@pytest.mark.parametrize("module", LOCK_MODULES)
def test_a_breaker_deletes_the_lockfile_it_judged(tmp_path, module):
    mod = importlib.import_module(module)
    lock = tmp_path / "write.lock"
    judged = _stale_record()
    lock.write_text(json.dumps(judged), encoding="utf-8")
    mod._break_lock(lock, judged)
    assert not lock.exists()
    assert not _guard_of(lock).exists(), "the guard file stayed"


@pytest.mark.parametrize("module", LOCK_MODULES)
def test_a_breaker_never_deletes_a_lockfile_published_since_it_looked(tmp_path, module):
    """The breaker read a dead holder's lockfile; by the time it breaks,
    another waiter has broken that one and published its own. The live lock
    must stay."""
    mod = importlib.import_module(module)
    lock = tmp_path / "write.lock"
    judged = _stale_record()
    live = {"pid": os.getpid(), "token": "a-live-holder", "time": time.time(),
            "host": locks._local_host()}
    lock.write_text(json.dumps(live), encoding="utf-8")
    mod._break_lock(lock, judged)
    assert lock.exists() and json.loads(lock.read_text(encoding="utf-8")) == live


@pytest.mark.parametrize("module", LOCK_MODULES)
def test_an_unreadable_lockfile_goes_only_while_it_is_still_unreadable_and_old(
        tmp_path, monkeypatch, module):
    mod = importlib.import_module(module)
    lock = tmp_path / "write.lock"
    lock.write_text("", encoding="utf-8")
    monkeypatch.setattr(mod, "_UNREADABLE_GRACE_SECONDS", 60.0)
    mod._break_lock(lock, {})
    assert lock.exists(), "a fresh unreadable lockfile was broken inside its grace"
    monkeypatch.setattr(mod, "_UNREADABLE_GRACE_SECONDS", -1.0)
    lock.write_text(json.dumps({"pid": os.getpid(), "token": "readable-now",
                                "time": time.time()}), encoding="utf-8")
    mod._break_lock(lock, {})
    assert lock.exists(), "a lockfile that can be read now was broken as unreadable"
    lock.write_text("not json", encoding="utf-8")
    mod._break_lock(lock, {})
    assert not lock.exists()


_GUARD_HOLDER = textwrap.dedent(
    """
    import importlib, sys, time
    from pathlib import Path
    sys.path.insert(0, sys.argv[2])
    lockrelease = importlib.import_module(sys.argv[4])
    with lockrelease.break_guard(Path(sys.argv[1])) as held:
        print("HELD" if held else "NOT", flush=True)
        time.sleep(float(sys.argv[3]))
    """
)


def test_the_guard_admits_one_breaker_at_a_time(tmp_path, monkeypatch):
    """Another process holds the guard: a breaker here waits, then leaves the
    lockfile alone, and its acquire loop looks again later."""
    lock = tmp_path / "write.lock"
    judged = _stale_record()
    lock.write_text(json.dumps(judged), encoding="utf-8")
    holder = subprocess.Popen(
        [sys.executable, "-c", _GUARD_HOLDER, str(lock), str(SRC), "3", lockrelease.__name__],
        stdout=subprocess.PIPE, text=True, stdin=subprocess.DEVNULL, creationflags=NO_WINDOW)
    try:
        assert holder.stdout.readline().strip() == "HELD"
        monkeypatch.setattr(lockrelease, "GUARD_WAIT_SECONDS", 0.3)
        with lockrelease.break_guard(lock) as held:
            assert held is False
        locks._break_lock(lock, judged)
        assert lock.exists(), "a breaker went ahead while another held the guard"
    finally:
        holder.wait(30)
    locks._break_lock(lock, judged)
    assert not lock.exists()


# ------------------------------------------------------ real processes


_STALE_RACER = textwrap.dedent(
    r"""
    import importlib, json, os, sys, time
    from pathlib import Path
    sys.path.insert(0, sys.argv[6])
    s = importlib.import_module(sys.argv[7])
    lock, slow, hold, out = Path(sys.argv[1]), float(sys.argv[3]), float(sys.argv[4]), sys.argv[5]
    if slow:
        alive = s._pid_alive
        def _slow(pid):
            r = alive(pid)
            time.sleep(slow)   # a breaker preempted between its read and its break
            return r
        s._pid_alive = _slow
    s._acquire_lockfile(lock, "book")
    t_in = time.time()
    time.sleep(hold)
    info = s._read_lock_info(lock)
    t_out = time.time()
    s._release_lockfile(lock)
    json.dump({"role": sys.argv[2], "t_in": t_in, "t_out": t_out,
               "still_mine": info.get("token") == s._OWNER_TOKEN}, open(out, "w"))
    """
)


def test_two_waiters_breaking_one_stale_lockfile_never_both_hold_it(tmp_path):
    """Made deterministic: a dead holder's lockfile is there. Waiter A reads
    it and is held up for a second before breaking it; waiter B breaks it at
    once, publishes its own and holds it. A must not delete B's live lock: A
    gets the lock only after B let it go, and B's lockfile is B's until
    then."""
    lock = tmp_path / "write.lock"
    lock.write_text(json.dumps(_stale_record()), encoding="utf-8")
    script = tmp_path / "racer.py"
    script.write_text(_STALE_RACER, encoding="utf-8")
    a_out, b_out = tmp_path / "a.json", tmp_path / "b.json"
    a = subprocess.Popen([sys.executable, str(script), str(lock), "A", "1.0", "0.2", str(a_out),
                          str(SRC), FILE_LOCK], stdin=subprocess.DEVNULL, creationflags=NO_WINDOW)
    time.sleep(0.3)
    b = subprocess.Popen([sys.executable, str(script), str(lock), "B", "0", "2.5", str(b_out),
                          str(SRC), FILE_LOCK], stdin=subprocess.DEVNULL, creationflags=NO_WINDOW)
    assert a.wait(60) == 0 and b.wait(60) == 0
    ra, rb = json.loads(a_out.read_text()), json.loads(b_out.read_text())
    assert not (ra["t_in"] < rb["t_out"] and rb["t_in"] < ra["t_out"]), (ra, rb)
    assert rb["still_mine"], "B's live lock was replaced while B held it"
    assert not lock.exists()


_STRESS_WORKER = textwrap.dedent(
    r"""
    import importlib, json, os, random, sys, time
    sys.path.insert(0, sys.argv[1])
    s = importlib.import_module(sys.argv[7])
    book, folder, name, rounds, seed = sys.argv[2], sys.argv[3], sys.argv[4], int(sys.argv[5]), int(sys.argv[6])
    rng = random.Random(seed)
    s.LOCK_WAIT_SECONDS = 120.0
    alive = s._pid_alive
    def _slow(pid):
        r = alive(pid)
        time.sleep(rng.choice((0, 0, 0.02, 0.08, 0.15)))   # preemption between read and break
        return r
    s._pid_alive = _slow
    counter = os.path.join(folder, "counter.txt")
    inside = os.path.join(folder, "inside-" + name)
    log = os.open(os.path.join(folder, "log-" + name), os.O_WRONLY | os.O_CREAT | os.O_APPEND)
    presence = os.open(os.path.join(folder, "presence"), os.O_RDWR | os.O_CREAT)

    def take_presence():
        # One holder at a time in the section, judged by an OS lock that the
        # system drops when a holder is killed. Held by a LIVE process while
        # this one also holds the write lock means two holders.
        end = time.monotonic() + 5
        while True:
            try:
                if sys.platform == "win32":
                    import msvcrt
                    os.lseek(presence, 0, 0)
                    msvcrt.locking(presence, msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(presence, fcntl.LOCK_EX | fcntl.LOCK_NB)
                os.lseek(presence, 8, 0)
                os.write(presence, str(os.getpid()).encode().ljust(12))
                return None
            except OSError:
                os.lseek(presence, 8, 0)
                other = os.read(presence, 12).strip()
                pid = int(other) if other.isdigit() else 0
                if pid and pid != os.getpid() and alive(pid):
                    return pid
                if time.monotonic() > end:
                    return -1
                time.sleep(0.005)

    def drop_presence():
        if sys.platform == "win32":
            import msvcrt
            os.lseek(presence, 0, 0)
            msvcrt.locking(presence, msvcrt.LK_UNLCK, 1)
        else:
            import fcntl
            fcntl.flock(presence, fcntl.LOCK_UN)

    overlaps = []
    for i in range(rounds):
        with s.write_lock(book):
            other = take_presence()
            if other is not None:
                overlaps.append(other)
            open(inside, "w").close()
            try:
                with open(counter) as f:
                    n = int(f.read() or 0)
            except FileNotFoundError:
                n = 0
            time.sleep(rng.uniform(0.005, 0.03))
            tmp = counter + "." + name
            with open(tmp, "w") as f:
                f.write(str(n + 1))
            os.replace(tmp, counter)
            os.write(log, b"+")
            os.remove(inside)
            if other is None:
                drop_presence()
        time.sleep(rng.uniform(0, 0.01))
    with open(os.path.join(folder, "done-" + name), "w") as f:
        json.dump({"overlaps": overlaps}, f)
    """
)


@pytest.mark.timeout(240)
def test_many_processes_with_stale_locks_and_killed_holders_never_two_holders(tmp_path):
    """Under load: six processes write one document's counter under the write
    lock. Breakers are held up at random between reading a lockfile and
    breaking it; a dead holder's lockfile is there at the start, and holders
    are killed in the middle of a hold, which leaves their lockfile behind
    for the others to break. No two processes are ever inside at once, and
    no increment is lost."""
    book = tmp_path / "book.dat"
    book.write_bytes(b"not a document; the lock never reads it")
    lock = locks.slot_dir(book, create=True) / locks.LOCK_FILE_NAME
    lock.write_text(json.dumps(_stale_record()), encoding="utf-8")
    folder = tmp_path / "run"
    folder.mkdir()
    script = tmp_path / "worker.py"
    script.write_text(_STRESS_WORKER, encoding="utf-8")
    workers = {}
    for i in range(6):
        name = f"w{i}"
        workers[name] = subprocess.Popen(
            [sys.executable, str(script), str(SRC), str(book), str(folder), name, "10",
             str(1000 + i), FILE_LOCK],
            stdin=subprocess.DEVNULL, creationflags=NO_WINDOW)
    killed = []
    end = time.monotonic() + 150
    while time.monotonic() < end and any(p.poll() is None for p in workers.values()):
        if len(killed) < 3:
            for name, proc in workers.items():
                if name not in killed and proc.poll() is None and (folder / f"inside-{name}").exists():
                    proc.kill()
                    proc.wait(30)
                    killed.append(name)
                    break
        time.sleep(0.15 if len(killed) < 3 else 0.5)
    for proc in workers.values():
        if proc.poll() is None:
            proc.kill()
            proc.wait(30)
            pytest.fail("a worker did not finish in time")
    assert len(killed) == 3, f"only {len(killed)} holders were killed mid-hold"
    overlaps = []
    for name in workers:
        done = folder / f"done-{name}"
        if name in killed:
            assert not done.exists()
            continue
        assert workers[name].returncode == 0, name
        overlaps += json.loads(done.read_text())["overlaps"]
    assert overlaps == [], f"two holders inside at once: {overlaps}"
    logged = sum(len((folder / f"log-{n}").read_bytes()) for n in workers
                 if (folder / f"log-{n}").exists())
    final = int((folder / "counter.txt").read_text())
    # A holder killed between its write and its log line counts once more.
    assert logged <= final <= logged + len(killed), (logged, final)
    assert logged >= 3 * 10, "the survivors did not finish their rounds"
    # The last lockfile can still be there: a holder killed after the others
    # finished, or a release whose sweeper thread had not deleted it when its
    # process exited. Either is stale, and the next writer breaks it: one more
    # hold, and then nothing is left.
    with locks.write_lock(book):
        assert lock.exists()
    assert not lock.exists(), "a lockfile was left behind"
    assert not _guard_of(lock).exists(), "the guard file stayed"
