"""DGS onayında dönem, proje ve şirket eşleşmesi için regresyon testleri."""
from contextlib import contextmanager
from contextlib import redirect_stdout
import io
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import types
import unittest
from unittest import mock

import dgs_onaya as O


class _NoSelectionPage:
    def locator(self, _selector):
        raise AssertionError("Yanlış dönemde portal seçimi yapılmamalı")


class _NodePage:
    def __init__(self, rows):
        self.rows = rows
        self.clicked = None

    def evaluate(self, expression, arg=None):
        script = """
          const rows = JSON.parse(process.argv[2]);
          const args = JSON.parse(process.argv[3]);
          let clicked = null;
          const trs = rows.map((cells, index) => ({
            children: cells.map(textContent => ({textContent})),
            scrollIntoView() {},
            dispatchEvent(event) { if (event.type === 'dblclick') clicked = index; }
          }));
          const body = {querySelectorAll(selector) { return selector === 'tr' ? trs : []; }};
          const dlg = {offsetParent: {}, querySelector(selector) {
            if (selector === '.ui-dialog-titlebar') return {textContent: 'Dışarıda Geçirilen Süreler Listesi'};
            if (selector === '.flexigrid .bDiv') return body;
            return null;
          }};
          global.document = {querySelectorAll(selector) { return selector === 'div.ui-dialog' ? [dlg] : []; }};
          global.window = {};
          global.MouseEvent = class {constructor(type) {this.type = type;}};
          const result = eval(process.argv[1])(args);
          console.log(JSON.stringify({result, clicked}));
        """
        result = subprocess.run(["node", "-e", script, expression,
                                 json.dumps(self.rows, ensure_ascii=False),
                                 json.dumps(arg, ensure_ascii=False)],
                                capture_output=True, text=True, check=True, timeout=10)
        payload = json.loads(result.stdout)
        self.clicked = payload["clicked"]
        return payload["result"]


def _row(project, person):
    cells = [""] * 16
    cells[0] = "1"
    cells[4] = project
    cells[6] = person + " 10*******18"
    cells[12] = "Değerlendirmeye Gönderilmemiş"
    cells[13] = "5"
    cells[15] = "10:00"
    return cells


class DgsOnayTests(unittest.TestCase):
    def test_missing_period_stops_without_selecting_another_month(self):
        options = [{"value": "101", "text": "EYLÜL 2026"},
                   {"value": "102", "text": "EKİM 2026"}]
        with mock.patch.object(O, "donem_options", return_value=options):
            with self.assertRaisesRegex(O.D.VerifyError, "AĞUSTOS 2026"):
                O.set_donem(_NoSelectionPage(), "AĞUSTOS 2026")

    @unittest.skipUnless(shutil.which("node"), "Node.js bulunamadı")
    def test_list_includes_dijitalpark_project_and_double_click_uses_project(self):
        page = _NodePage([_row("YANLIŞ PROJE", "ÖRNEK KİŞİ"),
                          _row("Etkin Süreç Projesi", "ÖRNEK KİŞİ")])
        rows = O.read_list_rows(page)
        self.assertEqual([r["project"] for r in rows],
                         ["YANLIŞ PROJE", "Etkin Süreç Projesi"])
        expected = {O.fold("ÖRNEK KİŞİ"): O.fold("Etkin Süreç Projesi")}
        self.assertEqual([r["project"] for r in rows if O._project_matches(r, expected)],
                         ["Etkin Süreç Projesi"])
        self.assertTrue(O.dblclick_row(page, O.fold("ÖRNEK KİŞİ"),
                                        O.fold("Etkin Süreç Projesi")))
        self.assertEqual(page.clicked, 1)

    def test_dijitalpark_commit_checks_company_before_opening_list(self):
        class CompanyMismatch(ValueError):
            pass

        tenant = types.ModuleType("portal_tenant")

        def check(_page, expected):
            self.assertEqual(expected, "ETKİN PROJE A.Ş.")
            raise CompanyMismatch("Yanlış şirket")

        tenant.CompanyMismatch = CompanyMismatch
        tenant.assert_expected_company = check
        tenant.resume_suffix = lambda expected: "companyhash12" if expected else None

        class Page:
            def on(self, *_args):
                pass

        @contextmanager
        def playwright():
            yield object()

        with tempfile.TemporaryDirectory() as folder:
            previous = os.getcwd()
            try:
                os.chdir(folder)
                Path("dgs_done_Dijitalpark_Eylül_companyhash12.txt").write_text(
                    "ÖRNEK KİŞİ\n", encoding="utf-8")
                person = O.D.PersonRow("ÖRNEK KİŞİ", "Etkin Süreç Projesi", "10:00",
                                       "Ar-Ge", "Dijitalpark")
                with mock.patch.object(sys, "argv", ["dgs_onaya.py", "--excel", "dummy.xlsx",
                                                    "--sheet", "Eylül", "--lokasyon", "Dijitalpark",
                                                    "--commit", "--donem", "EYLÜL 2026"]), \
                     mock.patch.object(O.D, "read_excel", return_value={"ÖRNEK KİŞİ": person}), \
                     mock.patch.object(O, "sync_playwright", playwright), \
                     mock.patch.object(O.D, "attach_browser", return_value=(object(), Page())), \
                     mock.patch.object(O.D, "assert_logged_in"), \
                     mock.patch.object(O, "open_list", side_effect=AssertionError("Liste açıldı")), \
                     mock.patch.dict(sys.modules, {"portal_tenant": tenant}), \
                     mock.patch.dict(os.environ, {"DGS_EXPECTED_COMPANY": "ETKİN PROJE A.Ş."}):
                    with self.assertRaisesRegex(O.D.VerifyError, "Yanlış şirket"):
                        O.main()
            finally:
                os.chdir(previous)

    def test_dijitalpark_dry_run_without_company_uses_excel_without_resume_file(self):
        class Page:
            def on(self, *_args):
                pass

        @contextmanager
        def playwright():
            yield object()

        tenant = types.ModuleType("portal_tenant")
        tenant.CompanyMismatch = type("CompanyMismatch", (ValueError,), {})
        tenant.assert_expected_company = mock.Mock(side_effect=AssertionError("Dry-run'da şirket kontrolü yok"))
        tenant.resume_suffix = mock.Mock(side_effect=AssertionError("Dry-run'da resume anahtarı yok"))
        person = O.D.PersonRow("ÖRNEK KİŞİ", "Etkin Süreç Projesi", "10:00",
                               "Ar-Ge", "Dijitalpark")
        with tempfile.TemporaryDirectory() as folder:
            previous = os.getcwd()
            try:
                os.chdir(folder)  # done dosyası bilerek yok
                with mock.patch.object(sys, "argv", ["dgs_onaya.py", "--excel", "dummy.xlsx",
                                                    "--sheet", "Eylül", "--lokasyon", "Dijitalpark",
                                                    "--donem", "EYLÜL 2026"]), \
                     mock.patch.object(O.D, "read_excel", return_value={"ÖRNEK KİŞİ": person}), \
                     mock.patch.object(O, "sync_playwright", playwright), \
                     mock.patch.object(O.D, "attach_browser", return_value=(object(), Page())), \
                     mock.patch.object(O.D, "assert_logged_in"), \
                     mock.patch.object(O, "open_list"), \
                     mock.patch.object(O, "set_donem"), \
                     mock.patch.object(O, "set_status_filter"), \
                     mock.patch.object(O, "listele"), \
                     mock.patch.object(O, "read_list_rows", return_value=[]), \
                     mock.patch.object(O, "footer_total", return_value="Toplam 0 kayıt"), \
                     mock.patch.dict(sys.modules, {"portal_tenant": tenant}), \
                     mock.patch.dict(os.environ, {"DGS_EXPECTED_COMPANY": ""}), \
                     mock.patch("builtins.input", side_effect=EOFError):
                    O.main()
            finally:
                os.chdir(previous)

    def test_dijitalpark_missing_draft_fails_reconciliation(self):
        class Page:
            def on(self, *_args):
                pass

            def evaluate(self, _javascript):
                return True

        @contextmanager
        def playwright():
            yield object()

        person = O.D.PersonRow("ÖRNEK KİŞİ", "Örnek Proje", "10:00",
                               "Ar-Ge", "Dijitalpark")
        with tempfile.TemporaryDirectory() as folder:
            previous = os.getcwd()
            try:
                os.chdir(folder)
                from portal_tenant import resume_suffix
                company = "ÖRNEK ŞİRKET ANONİM ŞİRKETİ"
                tag = resume_suffix(company)
                Path(f"dgs_done_Dijitalpark_Eylül_{tag}.txt").write_text(
                    "ÖRNEK KİŞİ\n", encoding="utf-8")
                with mock.patch.object(sys, "argv", ["dgs_onaya.py", "--excel", "dummy.xlsx",
                                                    "--sheet", "Eylül", "--lokasyon", "Dijitalpark",
                                                    "--commit", "--donem", "EYLÜL 2026"]), \
                     mock.patch.object(O.D, "read_excel", return_value={"ÖRNEK KİŞİ": person}), \
                     mock.patch.object(O, "sync_playwright", playwright), \
                     mock.patch.object(O.D, "attach_browser", return_value=(object(), Page())), \
                     mock.patch.object(O.D, "assert_logged_in"), \
                     mock.patch.object(O, "open_list"), \
                     mock.patch.object(O, "ensure_filter_open"), \
                     mock.patch.object(O, "set_donem"), \
                     mock.patch.object(O, "set_status_filter"), \
                     mock.patch.object(O, "listele"), \
                     mock.patch.object(O, "read_list_rows", return_value=[]), \
                     mock.patch("portal_tenant.assert_expected_company"), \
                     mock.patch.dict(os.environ, {"DGS_EXPECTED_COMPANY": company}), \
                     mock.patch("builtins.input", side_effect=EOFError):
                    with self.assertRaises(SystemExit) as raised:
                        O.main()
                self.assertEqual(raised.exception.code, 1)
            finally:
                os.chdir(previous)

    def test_legacy_park_keeps_fev_tr_filter_without_reading_excel(self):
        class Page:
            def on(self, *_args):
                pass

        @contextmanager
        def playwright():
            yield object()

        rows = [{"ad_fold": O.fold("ÖRNEK KİŞİ"), "project": "FEV TR",
                 "onay": "Değerlendirmeye Gönderilmemiş", "pers_raw": "ÖRNEK KİŞİ"},
                {"ad_fold": O.fold("ÖRNEK KİŞİ"), "project": "BAŞKA PROJE",
                 "onay": "Değerlendirmeye Gönderilmemiş", "pers_raw": "ÖRNEK KİŞİ"}]
        with tempfile.TemporaryDirectory() as folder:
            previous = os.getcwd()
            try:
                os.chdir(folder)
                Path("dgs_done_TPI_Eylül.txt").write_text("ÖRNEK KİŞİ\n", encoding="utf-8")
                output = io.StringIO()
                with mock.patch.object(sys, "argv", ["dgs_onaya.py", "--excel", "unused.xlsx",
                                                    "--sheet", "Eylül", "--lokasyon", "TPI",
                                                    "--donem", "EYLÜL 2026"]), \
                     mock.patch.object(O.D, "read_excel", side_effect=AssertionError("Eski park Excel okumaz")), \
                     mock.patch.object(O, "sync_playwright", playwright), \
                     mock.patch.object(O.D, "attach_browser", return_value=(object(), Page())), \
                     mock.patch.object(O.D, "assert_logged_in"), \
                     mock.patch.object(O, "open_list"), \
                     mock.patch.object(O, "set_donem"), \
                     mock.patch.object(O, "set_status_filter"), \
                     mock.patch.object(O, "listele"), \
                     mock.patch.object(O, "read_list_rows", return_value=rows), \
                     mock.patch.object(O, "footer_total", return_value="Toplam 2 kayıt"), \
                     mock.patch("builtins.input", side_effect=EOFError), \
                     redirect_stdout(output):
                    O.main()
                self.assertIn("Listede bu sayfada 1 kayıt", output.getvalue())
                self.assertIn("Bizim done hedeflerinden eşleşen: 1", output.getvalue())
            finally:
                os.chdir(previous)


if __name__ == "__main__":
    unittest.main()
