"""tools/check_public.py: what may not reach this public repository."""
import check_public


def test_repository_is_clean():
    assert check_public.scan_files() == []


def test_denied_names_are_found_without_being_listed():
    # The denylist holds digests only; these are the digests of the probes.
    probes = {check_public._digest(w) for w in ("probe-host", "probe.internal")}
    original = set(check_public.DENIED_DIGESTS)
    check_public.DENIED_DIGESTS |= probes
    try:
        assert check_public.denied_names("url = http://srv.probe.internal:8000/v1") == ["srv.probe.internal"]
        assert check_public.denied_names("provider: probe-host-swap") == ["probe-host-swap"]
        assert check_public.denied_names("PROBE-HOST_LOCAL_URL=1") == ["PROBE-HOST_LOCAL_URL"]
        assert check_public.denied_names("a probe hosting internals") == []
    finally:
        check_public.DENIED_DIGESTS.clear()
        check_public.DENIED_DIGESTS |= original


def test_other_checks():
    found = check_public.scan_text("x", "\n".join([
        "mail me: someone@company.example.org",
        "ok: a@example.com, b@users.noreply.github.com",
        "key = 0123456789abcdef0123456789abcdef",
        "host https://my-resource.openai.azure.com/openai/v1",
        "placeholder https://example-resource.openai.azure.com",
        "lan 10.1.2.3 and 192.168.0.10",
        "token apb_live_abcdef1234567890",
        "guid 12345678-1234-1234-1234-123456789abc",
        "graph 14d82eec-204b-4c2f-b7e8-296a70dab67e",
    ]))
    kinds = [line.split(": ")[1] for line in found]
    assert kinds.count("e-mail address") == 1
    assert "key-like hex" in kinds and "token prefix" in kinds and "GUID" in kinds
    assert kinds.count("Azure resource") == 1          # the placeholder is allowed
    assert kinds.count("private IP") == 2
    assert not any(":9:" in line for line in found)     # Microsoft's public Graph client id


def test_digest_helper_prints_a_line(capsys, monkeypatch):
    monkeypatch.setattr("sys.argv", ["check_public.py", "--digest", "Some-Name"])
    assert check_public.main() == 0
    assert capsys.readouterr().out.strip() == f'"{check_public._digest("some-name")}",'
