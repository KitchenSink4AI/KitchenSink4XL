"""Certificate-chained receipts, as the Business adapter verifies them.

WHY A CHAIN AT ALL. A flat receipt names a key id the build must already carry,
so retiring a signing key means rebuilding four products and reissuing every
outstanding receipt. A chained receipt carries the certificates that prove its
signing key was authorized under a root the build pins, and rotation then costs
one certificate. These tests are the behavioural spec for that, and they are
the product-side half of the 194-test suite in ks4-licensing: the same cases,
run against `ks4_license_business.chain` rather than against `ks4_chain`, so a
receipt the issuer accepts and a product refuses shows up here rather than at a
customer.

EVERY KEY IN THIS FILE IS A THROWAWAY, generated in the test and discarded with
it. The four REAL certificates and the REAL signed trust list are checked too,
in the section at the end, but only for what they actually are: public
material that must verify against the two real roots. No test reads a private
key, and nothing here can sign anything a customer would ever see.

THE PROPERTY UNDER TEST, above all others: windows and retirement cutoffs are
compared against the RECEIPT'S issued_on and never against the clock. A test
that passes today and fails in 2036 would be a test that proved the opposite,
so the windows here are deliberately in the past and in the future.
"""
import base64
import json
from pathlib import Path

import pytest
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

from ks4_license_business import chain, receipts
from ks4_license_business.policy import (BuildMetadataError, BusinessPolicy,
                                         ProductConfig, validate_build_metadata)

PRODUCT = "kitchensink4xl-business"
CONFIG = ProductConfig(product_id=PRODUCT,
                       display_name="KitchenSink4XL Business",
                       license_file_env="KS4XL_LICENSE_FILE")
TRUST = Path(__file__).resolve().parents[1] / "fixtures" / "trust"

#: Wide enough that no case in this file fails for a reason it did not mean to
#: test. Cases about windows name their own.
FAR = ("2020-01-01", "2035-01-01")
ISSUED = "2026-09-14"

WEB_SCOPE = {"name": "web", "bands": ["team", "department"],
             "products": ["kitchensink4word-business", "kitchensink4xl-business",
                          "kitchensink4ppt-business", "kitchensink4web-business",
                          "kitchensink4workflow-*", "ks4-workflow-*"]}
ENTERPRISE_SCOPE = {"name": "enterprise",
                    "bands": ["team", "department", "organization", "enterprise"],
                    "products": "*"}
ROOT_SCOPE = {"name": "root", "bands": "*", "products": "*"}


# --------------------------------------------------------------- throwaways


def mint(key_id: str) -> dict:
    """One throwaway Ed25519 keypair, in memory, never written to disk."""
    private = Ed25519PrivateKey.generate()
    public = private.public_key().public_bytes(serialization.Encoding.Raw,
                                               serialization.PublicFormat.Raw)
    return {"key_id": key_id, "private": private,
            "pub": base64.b64encode(public).decode("ascii")}


def certificate(signer: dict, subject: dict, cert_type: str, scope: dict,
                window=FAR) -> dict:
    """A signed certificate, built the way certify.py builds one.

    The claim is assembled here rather than imported from the issuing tools,
    because a test that borrowed the signer's own serializer would prove only
    that the serializer agrees with itself.
    """
    claim = {"schema": 1, "cert_type": cert_type,
             "issuer_key_id": signer["key_id"], "key_id": subject["key_id"],
             "public_key": subject["pub"], "valid_from": window[0],
             "valid_to": window[1], "scope": scope}
    return {"key_id": signer["key_id"], "claim": claim,
            "signature": base64.b64encode(
                signer["private"].sign(chain._canonical(claim))).decode("ascii")}


def receipt(key: dict, certs=None, *, product: str = PRODUCT,
            band: str | None = "team", issued_on: str = ISSUED,
            updates_through: str = "2027-09-14",
            licensee: str = "Chain Test Ltd",
            license_id: str = "CT-0001") -> bytes:
    payload = {"schema": 1, "product": product, "license_id": license_id,
               "licensee": licensee, "issued_on": issued_on,
               "updates_through": updates_through}
    if band is not None:
        payload["user_band"] = band
    envelope = {"key_id": key["key_id"], "payload": payload,
                "signature": base64.b64encode(
                    key["private"].sign(chain._canonical(payload))).decode("ascii")}
    if certs is not None:
        envelope["chain"] = certs
    return json.dumps(envelope).encode("utf-8")


def hierarchy(prefix: str = "t") -> dict:
    """Root, spare root, two intermediates, two signing keys, all certified.

    The shape the key ceremony produces, in miniature.
    """
    keys = {name: mint(f"{prefix}-{name}") for name in
            ("root", "root-spare", "int-web", "int-ent", "sign-web", "sign-ent")}
    certs = {
        "int-web": certificate(keys["root"], keys["int-web"], "intermediate",
                               WEB_SCOPE),
        "int-ent": certificate(keys["root"], keys["int-ent"], "intermediate",
                               ENTERPRISE_SCOPE),
    }
    certs["sign-web"] = certificate(keys["int-web"], keys["sign-web"], "signing",
                                    WEB_SCOPE)
    certs["sign-ent"] = certificate(keys["int-ent"], keys["sign-ent"], "signing",
                                    ENTERPRISE_SCOPE)
    return {"keys": keys, "certs": certs,
            "roots": {keys["root"]["key_id"]: keys["root"]["pub"],
                      keys["root-spare"]["key_id"]: keys["root-spare"]["pub"]}}


def chain_for(hier: dict, leaf: str = "sign-web") -> list:
    parent = "int-web" if leaf == "sign-web" else "int-ent"
    return [hier["certs"][leaf], hier["certs"][parent]]


def trust_list(hier: dict, version: int = 1, retirements=(),
               issued_on: str = ISSUED, extra_roots=()) -> dict:
    keys = hier["keys"]
    retired = {row["key_id"]: row.get("cutoff") for row in retirements}

    def entry(key_id, public_key, cross=None):
        row = {"key_id": key_id, "public_key": public_key,
               "status": "retired" if key_id in retired else "active",
               "retired_after": retired.get(key_id)}
        if cross is not None:
            row["cross_cert"] = cross
        return row

    roots = [entry(keys[n]["key_id"], keys[n]["pub"]) for n in ("root", "root-spare")]
    roots += [entry(k["key_id"], k["pub"], cross) for k, cross in extra_roots]
    return {"schema": 1, "version": version, "issued_on": issued_on,
            "roots": roots,
            "intermediates": [entry(keys[n]["key_id"], keys[n]["pub"])
                              for n in ("int-web", "int-ent")],
            "signing_keys": [entry(keys[n]["key_id"], keys[n]["pub"])
                             for n in ("sign-web", "sign-ent")],
            "retirements": [dict(row) for row in retirements]}


def sign_list(data: dict, signer: dict) -> dict:
    return {"key_id": signer["key_id"], "list": data,
            "signature": base64.b64encode(
                signer["private"].sign(chain._canonical(data))).decode("ascii")}


def build_meta(*, public_keys=None, roots=None, signed_list=None,
               accept_flat=None, release_date: str = ISSUED) -> dict:
    meta = {"edition": "business", "package": PRODUCT, "version": "2.1.3+business",
            "channel": "beta", "release_date": release_date,
            "public_keys": dict(public_keys or {})}
    if roots is not None:
        meta["root_keys"] = dict(roots)
    if signed_list is not None:
        meta["trust_list"] = signed_list
    if accept_flat is not None:
        meta["chain_policy"] = {"accept_flat": accept_flat}
    return meta


def refusal(raw: bytes, build: dict):
    with pytest.raises(receipts.ReceiptError) as problem:
        receipts.verify_receipt(raw, build, PRODUCT)
    return problem.value


# ------------------------------------------------- the two paths, both open


def test_a_chained_receipt_verifies_against_pinned_roots_alone():
    """The whole point: the build carries no signing key at all and still
    verifies a receipt, because the chain proves the signing key for it."""
    hier = hierarchy()
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier))
    out = receipts.verify_receipt(raw, build_meta(roots=hier["roots"]), PRODUCT)
    assert out["status"] == "verified"
    assert out["verification"] == "chain"
    assert out["licensed_to"] == "Chain Test Ltd"


def test_a_flat_receipt_still_verifies_on_a_chain_capable_build():
    """Compatibility for every receipt issued before certificates existed."""
    key = mint("flat-issuer")
    hier = hierarchy()
    raw = receipt(key)
    out = receipts.verify_receipt(
        raw, build_meta(public_keys={key["key_id"]: key["pub"]},
                        roots=hier["roots"]), PRODUCT)
    assert out["status"] == "verified"
    assert out["verification"] == "flat"


def test_a_chained_receipt_verifies_on_a_build_that_also_carries_the_flat_map():
    """Both maps present is the shipping state for this release."""
    hier = hierarchy()
    other = mint("flat-issuer")
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier))
    out = receipts.verify_receipt(
        raw, build_meta(public_keys={other["key_id"]: other["pub"]},
                        roots=hier["roots"]), PRODUCT)
    assert out["verification"] == "chain"


def test_a_chained_receipt_fits_under_the_shipped_size_cap():
    """A chain is carried inside the receipt, so it has to fit in the 16 KiB a
    shipped verifier will read. Two certificates leave plenty of room."""
    hier = hierarchy()
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier))
    assert len(raw) < receipts._MAX_LICENSE_BYTES


def test_the_single_cert_spelling_is_read_as_a_one_link_chain():
    """The legacy `cert` key, for a chain that stops at a root directly."""
    root = mint("solo-root")
    signer = mint("solo-sign")
    leaf = certificate(root, signer, "signing", ENTERPRISE_SCOPE)
    envelope = json.loads(receipt(signer))
    envelope["cert"] = leaf
    out = receipts.verify_receipt(json.dumps(envelope).encode(),
                                  build_meta(roots={root["key_id"]: root["pub"]}),
                                  PRODUCT)
    assert out["verification"] == "chain"


# ------------------------------------------------------------- the refusals


def test_a_chain_ending_at_an_unknown_root_is_refused():
    hier = hierarchy()
    stranger = mint("stranger-root")
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier))
    problem = refusal(raw, build_meta(roots={stranger["key_id"]: stranger["pub"]}))
    assert problem.category == "chain_invalid"
    assert problem.path == "chain"
    assert "not a trusted root" in problem.detail


def test_a_build_with_no_roots_refuses_a_chained_receipt():
    hier = hierarchy()
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier))
    key = hier["keys"]["sign-web"]
    # Even with the signing key sitting in the flat map: the receipt carries a
    # chain, so the chain is what is checked.
    problem = refusal(raw, build_meta(public_keys={key["key_id"]: key["pub"]}))
    assert problem.category == "chain_invalid"


def test_a_tampered_payload_on_a_chained_receipt_fails_on_the_signature():
    """The chain resolves, and then the receipt signature does not verify. The
    category is the signature one, because that is what went wrong."""
    hier = hierarchy()
    envelope = json.loads(receipt(hier["keys"]["sign-web"], chain_for(hier)))
    envelope["payload"]["licensee"] = "Somebody Else Ltd"
    problem = refusal(json.dumps(envelope).encode(),
                      build_meta(roots=hier["roots"]))
    assert problem.category == "signature_invalid"
    assert problem.path == "chain"


def test_a_tampered_certificate_claim_is_refused():
    hier = hierarchy()
    certs = chain_for(hier)
    certs[0] = json.loads(json.dumps(certs[0]))
    certs[0]["claim"]["scope"] = ENTERPRISE_SCOPE
    raw = receipt(hier["keys"]["sign-web"], certs)
    assert refusal(raw, build_meta(roots=hier["roots"])).category == "chain_invalid"


def test_a_missing_link_is_refused():
    """Leaf plus root cross-certificate, with the intermediate taken out."""
    hier = hierarchy()
    raw = receipt(hier["keys"]["sign-web"], [hier["certs"]["sign-web"]])
    problem = refusal(raw, build_meta(roots=hier["roots"]))
    assert problem.category == "chain_invalid"
    assert "not a trusted root" in problem.detail


def test_a_self_signed_leaf_is_refused():
    signer = mint("self-signed")
    leaf = certificate(signer, signer, "signing", ENTERPRISE_SCOPE)
    raw = receipt(signer, [leaf])
    root = mint("real-root")
    problem = refusal(raw, build_meta(roots={root["key_id"]: root["pub"]}))
    assert problem.category == "chain_invalid"
    assert "cannot certify itself" in problem.detail


def test_an_over_long_chain_is_refused():
    hier = hierarchy()
    certs = chain_for(hier) * 3
    raw = receipt(hier["keys"]["sign-web"], certs)
    problem = refusal(raw, build_meta(roots=hier["roots"]))
    assert "over the" in problem.detail


def test_a_repeated_key_id_in_the_chain_is_refused():
    hier = hierarchy()
    certs = chain_for(hier) + [hier["certs"]["sign-web"]]
    raw = receipt(hier["keys"]["sign-web"], certs)
    assert "twice" in refusal(raw, build_meta(roots=hier["roots"])).detail


def test_a_certificate_for_one_key_cannot_carry_another_keys_receipt():
    hier = hierarchy()
    raw = receipt(hier["keys"]["sign-ent"], chain_for(hier, "sign-web"))
    problem = refusal(raw, build_meta(roots=hier["roots"]))
    assert problem.category == "chain_invalid"


def test_an_intermediate_may_not_certify_an_intermediate():
    hier = hierarchy()
    keys = hier["keys"]
    rogue = mint("t-int-rogue")
    bad = certificate(keys["int-web"], rogue, "intermediate", WEB_SCOPE)
    signer = mint("t-sign-rogue")
    leaf = certificate(rogue, signer, "signing", WEB_SCOPE)
    raw = receipt(signer, [leaf, bad, hier["certs"]["int-web"]])
    problem = refusal(raw, build_meta(roots=hier["roots"]))
    assert "may not certify" in problem.detail


def test_a_signing_key_may_not_certify_anything():
    """The rule that stops a leaked leaf from minting its own hierarchy."""
    hier = hierarchy()
    victim = mint("t-minted")
    minted = certificate(hier["keys"]["sign-web"], victim, "signing", WEB_SCOPE)
    raw = receipt(victim, [minted, hier["certs"]["sign-web"],
                           hier["certs"]["int-web"]])
    assert "may not certify" in refusal(raw, build_meta(roots=hier["roots"])).detail


def test_an_envelope_key_id_may_not_disagree_with_the_signed_issuer():
    hier = hierarchy()
    certs = json.loads(json.dumps(chain_for(hier)))
    certs[0]["key_id"] = "somebody-else"
    raw = receipt(hier["keys"]["sign-web"], certs)
    assert "disagrees" in refusal(raw, build_meta(roots=hier["roots"])).detail


def test_an_intermediate_may_not_certify_beyond_its_own_scope():
    """Authority narrows downward. A web intermediate certifying an
    enterprise-scoped signing key is the rotation a stolen web key would want."""
    hier = hierarchy()
    over = mint("t-sign-over")
    leaf = certificate(hier["keys"]["int-web"], over, "signing", ENTERPRISE_SCOPE)
    raw = receipt(over, [leaf, hier["certs"]["int-web"]], band="team")
    problem = refusal(raw, build_meta(roots=hier["roots"]))
    assert "broader than its issuer grants" in problem.detail


# ------------------------------------------------------------------- scope


def test_a_web_key_cannot_sign_an_organization_band_receipt():
    """A stolen self-serve key cannot mint the bands it was never sold."""
    hier = hierarchy()
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier), band="organization")
    problem = refusal(raw, build_meta(roots=hier["roots"]))
    assert problem.category == "chain_invalid"
    assert "does not cover band organization" in problem.detail


def test_an_enterprise_key_may_sign_an_organization_band_receipt():
    hier = hierarchy()
    raw = receipt(hier["keys"]["sign-ent"], chain_for(hier, "sign-ent"),
                  band="organization")
    assert receipts.verify_receipt(raw, build_meta(roots=hier["roots"]),
                                   PRODUCT)["status"] == "verified"


def test_a_band_restricted_certificate_refuses_a_receipt_with_no_band():
    """Fails closed. The alternative is a web key signing anything at all
    simply by leaving the field out."""
    hier = hierarchy()
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier), band=None)
    assert "declares no band" in refusal(raw, build_meta(roots=hier["roots"])).detail


def test_a_certificate_scope_that_does_not_cover_the_product_is_refused():
    """The leaf is scoped to one product and the receipt names another.

    The scoped product has to satisfy two conditions at once, and getting
    either wrong makes this test pass for the wrong reason. It must be INSIDE
    the web intermediate's scope, or the chain is refused one step earlier for
    widening rather than at the product check. It must NOT be this repository's
    own product, and this file is copied into all four repositories with only
    the product id substituted, so naming any of the four by hand means a
    different thing in each copy. A workflow line satisfies both: the web scope
    covers it through the ks4-workflow-* pattern, and no editor is ever named
    that.
    """
    hier = hierarchy()
    scoped = {"name": "narrow", "bands": ["team"],
              "products": ["ks4-workflow-reporting"]}
    narrow = mint("t-sign-narrow")
    leaf = certificate(hier["keys"]["int-web"], narrow, "signing", scoped)
    raw = receipt(narrow, [leaf, hier["certs"]["int-web"]])
    problem = refusal(raw, build_meta(roots=hier["roots"]))
    assert "does not cover product" in problem.detail


def test_a_product_pattern_covers_a_line_that_does_not_exist_yet():
    hier = hierarchy()
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier),
                  product="ks4-workflow-reporting")
    out = receipts.verify_receipt(raw, build_meta(roots=hier["roots"]),
                                  "ks4-workflow-reporting")
    assert out["status"] == "verified"


def test_a_receipt_declaring_two_disagreeing_bands_is_refused():
    hier = hierarchy()
    envelope = json.loads(receipt(hier["keys"]["sign-web"], chain_for(hier)))
    envelope["payload"]["band"] = "department"
    # Re-sign, so the refusal is about the two fields and not the signature.
    envelope["signature"] = base64.b64encode(
        hier["keys"]["sign-web"]["private"].sign(
            chain._canonical(envelope["payload"]))).decode("ascii")
    problem = refusal(json.dumps(envelope).encode(), build_meta(roots=hier["roots"]))
    assert "different values" in problem.detail


# ---------------------------------------------------- windows, never a clock


def test_a_window_that_closed_years_ago_still_proves_a_receipt_dated_inside_it():
    """THE property. The certificate expired in 2023 by any clock and the
    receipt is dated 2022, so it verifies, today and in 2040."""
    root = mint("old-root")
    signer = mint("old-signer")
    leaf = certificate(root, signer, "signing", ENTERPRISE_SCOPE,
                       window=("2022-01-01", "2023-01-01"))
    raw = receipt(signer, [leaf], issued_on="2022-06-01",
                  updates_through="2023-06-01")
    out = receipts.verify_receipt(
        raw, build_meta(roots={root["key_id"]: root["pub"]},
                        release_date="2022-07-01"), PRODUCT)
    assert out["status"] == "verified"


def test_an_expired_intermediate_still_proves_a_receipt_dated_inside_it():
    root = mint("r")
    inter = mint("i")
    signer = mint("s")
    certs = [certificate(inter, signer, "signing", ENTERPRISE_SCOPE,
                         window=("2022-01-01", "2023-01-01")),
             certificate(root, inter, "intermediate", ENTERPRISE_SCOPE,
                         window=("2021-01-01", "2024-01-01"))]
    raw = receipt(signer, certs, issued_on="2022-06-01",
                  updates_through="2023-06-01")
    out = receipts.verify_receipt(
        raw, build_meta(roots={root["key_id"]: root["pub"]},
                        release_date="2022-07-01"), PRODUCT)
    assert out["status"] == "verified"


def test_a_receipt_dated_outside_the_certificate_window_is_refused():
    root = mint("r")
    signer = mint("s")
    leaf = certificate(root, signer, "signing", ENTERPRISE_SCOPE,
                       window=("2022-01-01", "2023-01-01"))
    raw = receipt(signer, [leaf], issued_on="2024-06-01")
    problem = refusal(raw, build_meta(roots={root["key_id"]: root["pub"]}))
    assert "outside the certificate window" in problem.detail


def test_the_chain_module_never_reads_the_clock():
    """Proved from the module's own syntax tree, because a test that exercised
    every path and missed one would still let a date.today() ship. Read as
    text rather than parsed, it would trip over the docstring that explains why
    the call is absent, which is why this walks the tree instead."""
    import ast

    tree = ast.parse(Path(chain.__file__).read_text(encoding="utf-8"))
    banned = {"today", "now", "utcnow", "fromtimestamp", "time", "monotonic"}
    found = [node.func.attr for node in ast.walk(tree)
             if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
             and node.func.attr in banned]
    assert found == [], f"the chain verifier reads the clock: {found}"


# ------------------------------------------------------------- retirements


def _retired_build(hier, key_name, cutoff):
    key_id = hier["keys"][key_name]["key_id"]
    rows = [{"key_id": key_id, "cutoff": cutoff, "reason": "rotation"}]
    data = trust_list(hier, retirements=rows)
    return build_meta(roots=hier["roots"],
                      signed_list=sign_list(data, hier["keys"]["root"]))


def test_a_retired_signing_keys_older_receipt_still_verifies():
    hier = hierarchy()
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier), issued_on="2026-06-01",
                  updates_through="2027-06-01")
    build = _retired_build(hier, "sign-web", "2026-08-01")
    assert receipts.verify_receipt(raw, build, PRODUCT)["status"] == "verified"


def test_a_retired_signing_keys_later_receipt_is_refused():
    hier = hierarchy()
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier), issued_on="2026-09-01")
    build = _retired_build(hier, "sign-web", "2026-08-01")
    problem = refusal(raw, build)
    assert problem.category == "chain_invalid"
    assert "was retired after 2026-08-01" in problem.detail


def test_an_unconditional_retirement_refuses_even_an_old_honest_receipt():
    """What a compromised key gets: no cutoff, so nothing under it verifies."""
    hier = hierarchy()
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier), issued_on="2021-01-05",
                  updates_through="2022-01-05")
    build = _retired_build(hier, "sign-web", None)
    assert "retired with no cutoff" in refusal(raw, build).detail


def test_retiring_an_intermediate_refuses_receipts_dated_after_the_cutoff():
    hier = hierarchy()
    build = _retired_build(hier, "int-web", "2026-08-01")
    later = receipt(hier["keys"]["sign-web"], chain_for(hier), issued_on="2026-09-01")
    assert refusal(later, build).category == "chain_invalid"
    earlier = receipt(hier["keys"]["sign-web"], chain_for(hier),
                      issued_on="2026-07-01", updates_through="2027-07-01")
    assert receipts.verify_receipt(earlier, build, PRODUCT)["status"] == "verified"


def test_retiring_a_root_refuses_receipts_dated_after_the_cutoff():
    hier = hierarchy()
    build = _retired_build(hier, "root", "2026-08-01")
    later = receipt(hier["keys"]["sign-web"], chain_for(hier), issued_on="2026-09-01")
    assert refusal(later, build).category == "chain_invalid"


def test_backdating_beats_a_date_based_retirement_and_that_is_the_known_limit():
    """DOCUMENTED, not fixed. A forged receipt backdated before a cutoff
    verifies on a product that has not been updated, which is the honest limit
    of offline licensing: cleanup rides the update channel and the ledger. The
    test exists so that nobody discovers the limit by accident later."""
    hier = hierarchy()
    build = _retired_build(hier, "sign-web", "2026-08-01")
    backdated = receipt(hier["keys"]["sign-web"], chain_for(hier),
                        issued_on="2026-01-01", updates_through="2027-01-01")
    assert receipts.verify_receipt(backdated, build, PRODUCT)["status"] == "verified"


# -------------------------------------------------------------- rotation


def test_rotating_a_signing_key_needs_no_rebuild():
    """The whole reason the hierarchy exists. The build is untouched; a new
    signing key gets a certificate under the same intermediate and its receipts
    verify, while the old key's receipts keep verifying too."""
    hier = hierarchy()
    build = build_meta(roots=hier["roots"])
    old = receipt(hier["keys"]["sign-web"], chain_for(hier))
    fresh = mint("t-sign-web-2027")
    cert = certificate(hier["keys"]["int-web"], fresh, "signing", WEB_SCOPE)
    new = receipt(fresh, [cert, hier["certs"]["int-web"]])
    assert receipts.verify_receipt(old, build, PRODUCT)["status"] == "verified"
    assert receipts.verify_receipt(new, build, PRODUCT)["status"] == "verified"


def test_the_spare_root_is_already_trusted_by_a_product_never_rebuilt():
    hier = hierarchy()
    spare = hier["keys"]["root-spare"]
    inter = mint("t-int-spare")
    signer = mint("t-sign-spare")
    certs = [certificate(inter, signer, "signing", ENTERPRISE_SCOPE),
             certificate(spare, inter, "intermediate", ENTERPRISE_SCOPE)]
    raw = receipt(signer, certs)
    assert receipts.verify_receipt(raw, build_meta(roots=hier["roots"]),
                                   PRODUCT)["status"] == "verified"


def test_a_cross_signed_root_is_reached_through_the_trust_list():
    """Root rotation on a product that pins only the old root: the trust list
    introduces the new root with a cross-certificate the old root signed, and
    a chain ending at the new root then verifies."""
    hier = hierarchy()
    fresh_root = mint("t-root-2029")
    cross = certificate(hier["keys"]["root"], fresh_root, "root", ROOT_SCOPE)
    data = trust_list(hier, extra_roots=[(fresh_root, cross)])
    signed = sign_list(data, hier["keys"]["root"])
    inter = mint("t-int-2029")
    signer = mint("t-sign-2029")
    certs = [certificate(inter, signer, "signing", ENTERPRISE_SCOPE),
             certificate(fresh_root, inter, "intermediate", ENTERPRISE_SCOPE)]
    raw = receipt(signer, certs)
    only_old = {hier["keys"]["root"]["key_id"]: hier["keys"]["root"]["pub"]}
    build = build_meta(roots=only_old, signed_list=signed)
    assert receipts.verify_receipt(raw, build, PRODUCT)["status"] == "verified"


def test_a_cross_certificate_from_a_key_the_build_does_not_pin_is_ignored():
    hier = hierarchy()
    stranger = mint("t-stranger")
    fresh_root = mint("t-root-forged")
    cross = certificate(stranger, fresh_root, "root", ROOT_SCOPE)
    data = trust_list(hier, extra_roots=[(fresh_root, cross)])
    signed = sign_list(data, hier["keys"]["root"])
    inter = mint("t-int-forged")
    signer = mint("t-sign-forged")
    certs = [certificate(inter, signer, "signing", ENTERPRISE_SCOPE),
             certificate(fresh_root, inter, "intermediate", ENTERPRISE_SCOPE)]
    raw = receipt(signer, certs)
    build = build_meta(roots=hier["roots"], signed_list=signed)
    assert "not a trusted root" in refusal(raw, build).detail


# ------------------------------------------------------------- trust lists


def test_a_trust_list_signed_by_a_key_the_build_does_not_pin_is_refused():
    hier = hierarchy()
    stranger = mint("t-stranger")
    signed = sign_list(trust_list(hier), stranger)
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier))
    assert refusal(raw, build_meta(roots=hier["roots"],
                                   signed_list=signed)).category == "chain_invalid"


def test_a_trust_list_signed_by_the_spare_root_is_accepted():
    hier = hierarchy()
    signed = sign_list(trust_list(hier), hier["keys"]["root-spare"])
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier))
    build = build_meta(roots=hier["roots"], signed_list=signed)
    assert receipts.verify_receipt(raw, build, PRODUCT)["status"] == "verified"


def test_a_trust_list_may_not_replace_the_public_key_of_a_pinned_root():
    hier = hierarchy()
    impostor = mint("x")
    data = trust_list(hier)
    data["roots"][0]["public_key"] = impostor["pub"]
    signed = sign_list(data, hier["keys"]["root-spare"])
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier))
    problem = refusal(raw, build_meta(roots=hier["roots"], signed_list=signed))
    assert "different public key than the one pinned" in problem.detail


def test_a_trust_list_version_may_only_move_forward():
    """The regression rule, checked directly: an attacker who keeps an old
    signed list cannot use it to un-retire a key a newer list pulled."""
    hier = hierarchy()
    current = trust_list(hier, version=4)
    older = sign_list(trust_list(hier, version=3), hier["keys"]["root"])
    with pytest.raises(chain.ChainError) as problem:
        chain.verify_trust_list(older, previous=current, roots=hier["roots"])
    assert "does not advance" in str(problem.value)
    newer = sign_list(trust_list(hier, version=5), hier["keys"]["root"])
    assert chain.verify_trust_list(newer, previous=current,
                                   roots=hier["roots"])["version"] == 5


def test_a_tampered_trust_list_body_is_refused():
    hier = hierarchy()
    signed = sign_list(trust_list(hier), hier["keys"]["root"])
    signed["list"]["issued_on"] = "2026-09-15"
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier))
    assert refusal(raw, build_meta(roots=hier["roots"],
                                   signed_list=signed)).category == "chain_invalid"


def test_a_retired_root_may_not_sign_the_list_that_succeeds_its_own():
    """A pulled root does not get to name its own successor list."""
    hier = hierarchy()
    root_id = hier["keys"]["root"]["key_id"]
    current = trust_list(hier, version=1,
                         retirements=[{"key_id": root_id, "cutoff": "2026-08-01"}])
    nxt = sign_list(trust_list(hier, version=2), hier["keys"]["root"])
    with pytest.raises(chain.ChainError) as problem:
        chain.verify_trust_list(nxt, previous=current)
    assert "not among" in str(problem.value)


# -------------------------------------------------- the build-time switch


def test_require_chain_refuses_a_flat_receipt():
    key = mint("flat-issuer")
    hier = hierarchy()
    build = build_meta(public_keys={key["key_id"]: key["pub"]},
                       roots=hier["roots"], accept_flat=False)
    problem = refusal(receipt(key), build)
    assert problem.category == "chain_invalid"
    assert problem.path == "flat"
    assert "reissued receipt" in str(problem)


def test_require_chain_still_verifies_a_chained_receipt():
    hier = hierarchy()
    build = build_meta(roots=hier["roots"], accept_flat=False)
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier))
    assert receipts.verify_receipt(raw, build, PRODUCT)["verification"] == "chain"


def test_a_build_that_says_nothing_about_chain_policy_accepts_flat_receipts():
    """Every _edition.json generated before this field existed."""
    key = mint("flat-issuer")
    build = {"edition": "business", "package": PRODUCT, "version": "2.1.3+business",
             "release_date": ISSUED, "public_keys": {key["key_id"]: key["pub"]}}
    assert receipts.accepts_flat(build) is True
    assert receipts.verify_receipt(receipt(key), build, PRODUCT)["status"] == "verified"


# ------------------------------------------------------- the build metadata


def test_build_metadata_carrying_chain_material_validates():
    hier = hierarchy()
    signed = sign_list(trust_list(hier), hier["keys"]["root"])
    meta = build_meta(roots=hier["roots"], signed_list=signed, accept_flat=True)
    assert validate_build_metadata(meta, CONFIG) is meta


def test_build_metadata_with_an_unusable_root_is_a_setup_error():
    meta = build_meta(roots={"ks4-root-2026": "not base64 at all"})
    with pytest.raises(BuildMetadataError):
        validate_build_metadata(meta, CONFIG)


def test_build_metadata_with_a_root_id_unfit_for_a_name_is_a_setup_error():
    root = mint("ok")
    meta = build_meta(roots={"../escape": root["pub"]})
    with pytest.raises(BuildMetadataError):
        validate_build_metadata(meta, CONFIG)


def test_build_metadata_with_a_trust_list_its_roots_do_not_vouch_for_is_refused():
    hier = hierarchy()
    stranger = mint("t-stranger")
    meta = build_meta(roots=hier["roots"],
                      signed_list=sign_list(trust_list(hier), stranger))
    with pytest.raises(BuildMetadataError):
        validate_build_metadata(meta, CONFIG)


def test_build_metadata_with_a_trust_list_and_no_roots_is_refused():
    hier = hierarchy()
    meta = build_meta(signed_list=sign_list(trust_list(hier), hier["keys"]["root"]))
    with pytest.raises(BuildMetadataError):
        validate_build_metadata(meta, CONFIG)


def test_build_metadata_with_an_unusable_chain_policy_is_refused():
    hier = hierarchy()
    meta = build_meta(roots=hier["roots"])
    meta["chain_policy"] = {"accept_flat": "yes please"}
    with pytest.raises(BuildMetadataError):
        validate_build_metadata(meta, CONFIG)


def test_unknown_build_metadata_fields_are_still_tolerated():
    """A newer builder writing a field this adapter does not read has to keep
    working, which is why extras are not an error."""
    hier = hierarchy()
    meta = build_meta(roots=hier["roots"])
    meta["something_the_next_release_adds"] = {"any": "shape"}
    assert validate_build_metadata(meta, CONFIG) is meta


# ------------------------------------------------------- what the server says


def _policy(tmp_path, monkeypatch, build, raw=None):
    if raw is not None:
        path = tmp_path / "receipt.json"
        path.write_bytes(raw)
        monkeypatch.setenv("KS4XL_LICENSE_FILE", str(path))
    else:
        monkeypatch.delenv("KS4XL_LICENSE_FILE", raising=False)
    return BusinessPolicy(CONFIG, build)


def test_the_policy_reports_the_chain_facts(tmp_path, monkeypatch):
    hier = hierarchy()
    signed = sign_list(trust_list(hier, version=1), hier["keys"]["root"])
    policy = _policy(tmp_path, monkeypatch,
                     build_meta(roots=hier["roots"], signed_list=signed))
    assert policy.chain_version == 1
    assert policy.trust_list_version == 1
    assert policy.root_key_ids == sorted(hier["roots"])
    assert policy.accepts_flat_receipts is True


def test_a_build_with_no_roots_reports_no_chain_version(tmp_path, monkeypatch):
    key = mint("flat-issuer")
    policy = _policy(tmp_path, monkeypatch,
                     build_meta(public_keys={key["key_id"]: key["pub"]}))
    assert policy.chain_version is None
    assert policy.trust_list_version is None
    assert policy.root_key_ids == []


def test_describe_license_names_the_path_that_verified(tmp_path, monkeypatch):
    hier = hierarchy()
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier))
    policy = _policy(tmp_path, monkeypatch, build_meta(roots=hier["roots"]), raw)
    out = policy.describe_license()
    assert out["status"] == "verified"
    assert out["verification"] == "chain"
    assert "licensed_to" not in out


def test_describe_license_names_the_path_that_was_refused(tmp_path, monkeypatch):
    hier = hierarchy()
    stranger = mint("stranger-root")
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier))
    policy = _policy(tmp_path, monkeypatch,
                     build_meta(roots={stranger["key_id"]: stranger["pub"]}), raw)
    out = policy.describe_license()
    assert out["status"] == "unverified"
    assert out["reason"] == "chain_invalid"
    assert out["verification"] == "chain"


def test_a_refusal_never_leaks_the_sub_reason_to_the_caller(tmp_path, monkeypatch):
    """The category is a closed set a caller may branch on. The verifier's own
    words can quote a key id or a band out of the receipt, so they go to a
    debug log and nowhere near the returned dict."""
    hier = hierarchy()
    raw = receipt(hier["keys"]["sign-web"], chain_for(hier), band="organization")
    policy = _policy(tmp_path, monkeypatch, build_meta(roots=hier["roots"]), raw)
    out = policy.describe_license(redacted=False)
    assert out["reason"] == "chain_invalid"
    assert "organization" not in json.dumps(out)
    assert "Chain Test Ltd" not in json.dumps(out)


def test_an_unconfigured_installation_still_reports_offline(tmp_path, monkeypatch):
    hier = hierarchy()
    out = _policy(tmp_path, monkeypatch, build_meta(roots=hier["roots"])).describe_license()
    assert out["status"] == "missing"
    assert out["verification"] == "offline"


# ------------------------------------------- the real, public trust material


def real_roots() -> dict:
    return json.loads((TRUST / "roots.json").read_text(encoding="utf-8"))


def real_cert(name: str) -> dict:
    return json.loads((TRUST / "certs" / f"{name}.cert.json").read_text(
        encoding="utf-8"))


def real_trust_list() -> dict:
    return json.loads((TRUST / "trustlist-v1.signed.json").read_text(
        encoding="utf-8"))


#: The canonical fingerprints from ks4-licensing trust/README.md, which is the
#: form read aloud when comparing a key out of band. They are written out here
#: rather than computed, because a test that computed both sides would prove
#: only that sha256 is deterministic.
REAL_ROOT_FINGERPRINTS = {
    "ks4-root-2026":
        "sha256:db2b088e2f28deaae39107b2b91f98d5becf136b673e36fdeadc3f8eac713229",
    "ks4-root-2026b":
        "sha256:6948d3baca358a48c59cfbfb57ee40aa463eaa99a213f9d7977a3fd2a92962a3",
}


def test_the_real_roots_are_the_published_ones():
    roots = real_roots()
    assert sorted(roots) == sorted(REAL_ROOT_FINGERPRINTS)
    for key_id, b64 in roots.items():
        assert chain.fingerprint(b64) == REAL_ROOT_FINGERPRINTS[key_id]


@pytest.mark.parametrize("name,issuer", [
    ("ks4-int-web-2026", "ks4-root-2026"),
    ("ks4-int-ent-2026", "ks4-root-2026"),
])
def test_the_real_intermediates_verify_under_the_real_root(name, issuer):
    cert = chain.validate_cert(real_cert(name), where=name)
    assert cert["claim"]["cert_type"] == "intermediate"
    assert cert["claim"]["issuer_key_id"] == issuer
    chain.check_cert_signature(cert, real_roots()[issuer], where=name)


@pytest.mark.parametrize("name,issuer", [
    ("ks4-issuer-web-2026", "ks4-int-web-2026"),
    ("ks4-issuer-2026", "ks4-int-ent-2026"),
])
def test_the_real_signing_certificates_verify_under_their_intermediates(name, issuer):
    cert = chain.validate_cert(real_cert(name), where=name)
    parent = chain.validate_cert(real_cert(issuer), where=issuer)
    assert cert["claim"]["cert_type"] == "signing"
    assert cert["claim"]["issuer_key_id"] == issuer
    chain.check_cert_signature(cert, parent["claim"]["public_key"], where=name)


def test_the_real_signing_scopes_never_widen_on_their_intermediates():
    for leaf, parent in (("ks4-issuer-web-2026", "ks4-int-web-2026"),
                         ("ks4-issuer-2026", "ks4-int-ent-2026")):
        child = real_cert(leaf)["claim"]["scope"]
        above = real_cert(parent)["claim"]["scope"]
        assert chain.scope_covers(above, child), chain.scope_gap(above, child)


def test_the_real_web_scope_sells_only_the_self_serve_bands():
    scope = real_cert("ks4-issuer-web-2026")["claim"]["scope"]
    assert scope["bands"] == ["team", "department"]


def test_the_real_trust_list_verifies_under_the_real_roots():
    body = chain.verify_trust_list(real_trust_list(), roots=real_roots())
    assert body["version"] == 1
    assert body["retirements"] == []
    assert {e["key_id"] for e in body["roots"]} == set(real_roots())


def test_the_real_trust_list_carries_every_real_certificate_subject():
    body = chain.verify_trust_list(real_trust_list(), roots=real_roots())
    listed = {e["key_id"] for section in chain.SECTIONS for e in body[section]}
    for name in ("ks4-int-web-2026", "ks4-int-ent-2026", "ks4-issuer-web-2026",
                 "ks4-issuer-2026"):
        claim = real_cert(name)["claim"]
        assert claim["key_id"] in listed
        entry = next(e for section in chain.SECTIONS for e in body[section]
                     if e["key_id"] == claim["key_id"])
        # The list and the certificate must name the same key, or a retirement
        # would be aimed at something other than what the chain carries.
        assert entry["public_key"] == claim["public_key"]


def test_the_real_trust_list_does_not_verify_under_a_stranger():
    stranger = mint("not-a-root")
    with pytest.raises(chain.ChainError):
        chain.verify_trust_list(real_trust_list(),
                                roots={stranger["key_id"]: stranger["pub"]})


def test_a_build_carrying_the_real_roots_and_list_is_valid_metadata():
    """What the shipped Business edition carries, checked as an artifact would
    be checked at install time. No receipt is involved: the real signing keys
    are offline and this test signs nothing."""
    meta = build_meta(roots=real_roots(), signed_list=real_trust_list())
    assert validate_build_metadata(meta, CONFIG) is meta
    policy = BusinessPolicy(CONFIG, meta)
    assert policy.root_key_ids == sorted(REAL_ROOT_FINGERPRINTS)
    assert policy.trust_list_version == 1
    assert policy.chain_version == 1
