"""The honest notice for clients that cannot see newly enabled tools.

Some MCP clients fix their tool list when the session or the worker starts.
Claude Code workers and the Codex CLI both do. For those, a pack switched on
by enable_tools mid-session never becomes callable, while enable_tools keeps
answering that the pack is on. An agent in that position retries the call, or
worse, reports the capability as missing from the server.

So the server says it. Three places, one fact:

  - every successful enable_tools result, INCLUDING the no-op re-enable that
    a stuck agent reaches on its second try, which used to come back with no
    note at all;
  - the server instructions, read once at the handshake, before any tool call;
  - the get_workflows index, read by an agent planning a multi-step job.

Nothing was built for this. The route out already existed: KS4XL_MODE takes a
comma list of packs at launch (packs.apply_startup_mode), so the notice points
at that instead of promising a refresh the client will not perform.

These tests pin the wording and the emission, because both are the whole fix:
a notice that is not emitted on the second call, or that has drifted between
its two prose copies, is the defect back again.
"""

from __future__ import annotations

import asyncio

import pytest

from xlsx_mcp import packs, server
from xlsx_mcp.ops import workflows as _workflows


@pytest.fixture
def restore_enabled():
    """Snapshot and restore the process-global enabled bookkeeping."""
    saved = dict(packs._ENABLED)
    yield
    packs._ENABLED.clear()
    packs._ENABLED.update(saved)
    server._PENDING_VISIBILITY.clear()


def _reset_to_lite():
    for name in packs._ENABLED:
        packs._ENABLED[name] = packs.pack_of(name) == "lite"


# --------------------------------------------- (a) the enable_tools note


def test_enable_note_states_what_a_fixed_tool_list_means(restore_enabled):
    """A first enable carries the note, and the note carries the four things
    an agent needs: what happened, the route that works anywhere, the one
    that works in Claude Code, and that a refusal is not worth retrying."""
    _reset_to_lite()
    result = packs.enable(["design"])
    assert result["enabled"] == ["design"]
    note = result["note"]
    assert note == packs.client_note(True)
    assert note.startswith("tools/list_changed was sent.")
    assert "this client fixed its list when the session or worker started" \
        in note
    assert "do not retry here" in note
    assert "What works in every client:" in note
    assert "add the packs to KS4XL_MODE (comma list)" in note
    assert "restart the app or session" in note
    assert "Claude Code only:" in note
    assert "call enable_tools in the main session" in note
    assert "start a new worker" in note
    assert "an administrator locked the tool set: do not retry" in note


def test_enable_note_leads_with_the_route_that_works_anywhere():
    """The order is the finding, not a style choice. The start-up route was
    proven end to end in Codex CLI including workers; the orchestrator route
    works in Claude Code and does NOT reach a Codex worker, so it comes
    second and carries its client's name. Leading with it would send most
    readers down a path that cannot help them."""
    note = packs.CLIENT_NOTE_BODY
    everywhere = note.index("What works in every client:")
    claude_code = note.index("Claude Code only:")
    assert everywhere < claude_code, \
        "the universal route must come before the client-specific one"
    assert note.index("KS4XL_MODE") < claude_code, \
        "the launch env belongs to the universal route, not the second one"


def test_note_rides_the_no_op_re_enable(restore_enabled):
    """THE regression. The note used to be gated on something having flipped,
    so the second identical call, which is exactly where an agent whose client
    ignored the first one lands, returned no note at all.

    The advice is the same on both calls; only the first sentence differs,
    because only the first call actually notified anything."""
    _reset_to_lite()
    first = packs.enable(["design"])
    second = packs.enable(["design"])
    assert second["enabled"] == []
    assert second["already_enabled"] == ["design"]
    assert second["approx_tokens_added"] == 0
    assert second["note"] == packs.client_note(False)
    assert second["note"] != first["note"]
    assert packs.CLIENT_NOTE_BODY in second["note"]


def test_no_op_prefix_does_not_claim_a_notification(restore_enabled):
    """A no-op sends nothing, so the note must not say it sent something.
    An agent debugging a tool it cannot see would otherwise go hunting for a
    notification that was never emitted."""
    _reset_to_lite()
    packs.enable(["design"])
    note = packs.enable(["design"])["note"]
    assert note.startswith(
        "These packs were already on, so no list change was sent.")
    assert "tools/list_changed was sent." not in note


def test_no_op_really_emits_nothing(restore_enabled):
    """The claim behind the prefix, proven rather than assumed: _sync gates on
    a non-empty name set, so the no-op never reaches the visibility hook and
    nothing is queued for the session."""
    _reset_to_lite()
    packs.enable(["design"])
    server._PENDING_VISIBILITY.clear()
    packs.enable(["design"])
    assert server._PENDING_VISIBILITY == []


def test_note_rides_a_mixed_call(restore_enabled):
    """One pack already on, one not. Something DID flip, so this is a real
    list change and takes the sent prefix."""
    _reset_to_lite()
    packs.enable(["design"])
    mixed = packs.enable(["design", "io"])
    assert mixed["enabled"] == ["io"]
    assert mixed["already_enabled"] == ["design"]
    assert mixed["note"] == packs.client_note(True)


def test_note_no_longer_promises_a_refresh(restore_enabled):
    """The retired sentence assumed the client could re-fetch. Several cannot,
    and telling them to was the whole defect."""
    _reset_to_lite()
    note = packs.enable(["design"])["note"]
    assert "re-fetch the tool list" not in note
    assert "does not refresh automatically" not in note


def test_note_survives_the_server_layer(restore_enabled):
    """enable_tools is a thin async wrapper over packs.enable; the note has to
    reach the caller through it, on the first call and on the no-op."""
    _reset_to_lite()

    async def run():
        first = await server.enable_tools(["io"])
        second = await server.enable_tools(["io"])
        return first, second

    first, second = asyncio.run(run())
    assert first["note"] == packs.client_note(True)
    assert second["enabled"] == []
    assert second["note"] == packs.client_note(False)


def test_disable_carries_no_such_note(restore_enabled):
    """Losing tools mid-session is not the failure mode being described; a
    client that never adds tools still drops them. No note, no noise."""
    _reset_to_lite()
    packs.enable(["design"])
    assert "note" not in packs.disable(["design"])


# ------------------------------- (b) the fact, before the first tool call


def test_instructions_carry_the_worker_surface_note():
    """Read once at the handshake, so a worker learns its surface is fixed
    before it spends a call finding out."""
    text = server.mcp.instructions or ""
    assert packs.WORKER_SURFACE_NOTE in text


def test_get_workflows_index_carries_the_worker_surface_note():
    """An agent planning a multi-step job reads the index; the pack advice it
    finds there is incomplete without this."""
    note = _workflows.get_workflows()["note"]
    assert packs.WORKER_SURFACE_NOTE in note
    assert "each step names its tool" in note


def test_worker_surface_note_says_both_routes_in_the_same_order():
    """Set the launch env, or, in Claude Code, enable before starting
    workers. Naming only one leaves half the users stuck, and the same
    ordering finding applies here as in the enable_tools note: the route
    that works anywhere leads."""
    note = packs.WORKER_SURFACE_NOTE
    assert "only see the tools that were on when they started" in note
    assert "KS4XL_MODE set to a comma list of packs" in note
    assert "enable packs in the main session before starting workers" in note
    assert note.index("KS4XL_MODE") < note.index("in Claude Code")


def test_worker_surface_note_is_single_sourced():
    """Two prose copies of one sentence drift. These are the same object."""
    assert packs.WORKER_SURFACE_NOTE in (server.mcp.instructions or "")
    assert packs.WORKER_SURFACE_NOTE in _workflows.get_workflows()["note"]


def test_the_advertised_env_really_takes_a_comma_list(restore_enabled,
                                                      monkeypatch):
    """The notice sends the user to KS4XL_MODE with a comma list. If that
    stopped working the advice would be worse than none."""
    monkeypatch.setenv("KS4XL_MODE", "design,io")
    _reset_to_lite()
    packs.apply_startup_mode()
    for pack in ("design", "io"):
        for name in packs.pack_tools(pack):
            assert packs.is_tool_enabled(name), name


# ------------------------------------------------- (c) the locked refusal


def test_locked_refusal_already_names_the_human_who_set_it(restore_enabled,
                                                           monkeypatch):
    """The notice tells an agent that a refused pack was turned off by a
    person, so the refusal itself must say the same. It already did, which is
    why no sentence was added to it; this pins that so a reword cannot quietly
    turn the refusal back into an unexplained CONFLICT."""
    monkeypatch.setenv("KS4XL_PACK_POLICY", "locked")
    with pytest.raises(Exception) as exc:
        packs.enable(["design"])
    message = str(exc.value)
    assert "fixed at startup by the host" in message
    assert "Ask the operator" in message


# ------------------------------------------------------------- the gates


@pytest.mark.parametrize("text", [
    packs.client_note(True),
    packs.client_note(False),
    packs.WORKER_SURFACE_NOTE,
])
def test_notice_strings_carry_no_em_dash(text):
    """The repo bans em dashes in everything a user reads."""
    assert "—" not in text
