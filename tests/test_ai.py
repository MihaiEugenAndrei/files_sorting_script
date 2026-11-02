from types import SimpleNamespace

from sort_files import MAX_TOKENS_RETRY, ask_local_ai, build_final_name, parse_ai_json


def make_client(responses):
    """responses: listă de string-uri (raspuns) sau Exception (eroare)."""
    calls = []

    def create(**kwargs):
        calls.append(kwargs)
        r = responses.pop(0)
        if isinstance(r, Exception):
            raise r
        msg = SimpleNamespace(content=r)
        return SimpleNamespace(choices=[SimpleNamespace(message=msg)])

    completions = SimpleNamespace(create=create)
    chat = SimpleNamespace(completions=completions)
    client = SimpleNamespace(chat=chat, calls=calls)
    return client


class TestParseAiJson:
    def test_plain_json(self):
        data = parse_ai_json('{"categorie": "Financiar", "subiect": "Facturi", "nume_nou": "Factura ING", "pastreaza_original": false}')
        assert data["categorie"] == "Financiar"

    def test_fenced_json(self):
        raw = '```json\n{"categorie": "Financiar", "subiect": "Facturi", "nume_nou": "X", "pastreaza_original": true}\n```'
        assert parse_ai_json(raw)["subiect"] == "Facturi"

    def test_prose_around_json(self):
        raw = 'Iată rezultatul: {"categorie": "Unknown", "subiect": "", "nume_nou": "a", "pastreaza_original": true} sper că ajută'
        assert parse_ai_json(raw)["categorie"] == "Unknown"

    def test_truncated_json_returns_none(self):
        assert parse_ai_json('{"categorie": "Financiar", "subiec') is None

    def test_missing_categorie_returns_none(self):
        assert parse_ai_json('{"subiect": "Facturi"}') is None

    def test_empty(self):
        assert parse_ai_json("") is None
        assert parse_ai_json(None) is None


class TestAskLocalAi:
    def test_valid_first_try(self):
        client = make_client(['{"categorie": "Financiar", "subiect": "Facturi", "nume_nou": "Factura Raiffeisen", "pastreaza_original": false}'])
        res = ask_local_ai(client, "fake", "text", "payload", "scan.pdf")
        assert res["ai_ok"] is True
        assert res["categorie"] == "Financiar"
        assert len(client.calls) == 1
        assert client.calls[0]["max_tokens"] == 90
        assert client.calls[0]["temperature"] == 0.0

    def test_retry_with_more_tokens_on_truncated_json(self):
        client = make_client([
            '{"categorie": "Financiar", "subiec',  # trunchiat
            '{"categorie": "Financiar", "subiect": "Facturi", "nume_nou": "Factura ING", "pastreaza_original": false}',
        ])
        res = ask_local_ai(client, "fake", "text", "payload", "scan.pdf")
        assert res["ai_ok"] is True
        assert len(client.calls) == 2
        assert client.calls[1]["max_tokens"] == MAX_TOKENS_RETRY

    def test_total_failure_returns_fallback(self):
        client = make_client(["gunoi", "și mai mult gunoi"])
        res = ask_local_ai(client, "fake", "text", "payload", "scan001.pdf")
        assert res["ai_ok"] is False
        assert res["categorie"] == "Unknown"
        assert res["pastreaza_original"] is True
        assert res["nume_nou"] == "scan001.pdf"

    def test_exception_returns_fallback(self):
        client = make_client([ConnectionError("serverul nu răspunde")])
        res = ask_local_ai(client, "fake", "text", "payload", "a.txt")
        assert res["ai_ok"] is False

    def test_image_mode_sends_image_url(self):
        client = make_client(['{"categorie": "Fotografii și Grafică", "subiect": "Imagini", "nume_nou": "Poza", "pastreaza_original": false}'])
        res = ask_local_ai(client, "fake", "image", "QUJD", "poza.jpg")
        assert res["ai_ok"] is True
        content = client.calls[0]["messages"][1]["content"]
        assert isinstance(content, list)
        assert content[1]["type"] == "image_url"
        assert content[1]["image_url"]["url"].startswith("data:image/jpeg;base64,")


class TestBuildFinalName:
    def test_renamed_appends_original_stem(self):
        res = {"nume_nou": "Factura Raiffeisen Utilitati", "pastreaza_original": False}
        assert build_final_name(res, "scan_001", ".pdf") == "Factura Raiffeisen Utilitati (scan 001).pdf"

    def test_keep_original(self):
        res = {"nume_nou": "whatever", "pastreaza_original": True}
        assert build_final_name(res, "contract final", ".docx") == "contract final.docx"

    def test_same_name_keeps_original_only(self):
        res = {"nume_nou": "Raport 2026", "pastreaza_original": False}
        assert build_final_name(res, "raport_2026", ".pdf") == "raport 2026.pdf"

    def test_empty_nume_nou(self):
        res = {"nume_nou": "", "pastreaza_original": False}
        assert build_final_name(res, "fisier", ".txt") == "fisier.txt"
