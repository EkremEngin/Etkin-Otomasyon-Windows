"""Kapanış tekrarlarında ana girişin onay ve personel kapsamı korunur."""
import contextlib
import io
import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import dgs_park


class KapanisTests(unittest.TestCase):
    def test_gui_command_keeps_operator_choices(self):
        base = ["worker", "dgs"]
        draft = dgs_park.kapanis_komutu(base, "Dijitalpark", "draft.xlsx",
                                         onayla=False, destek=False)
        self.assertIn("--commit", draft)
        self.assertNotIn("--onayla", draft)
        self.assertNotIn("--include-destek", draft)
        approved = dgs_park.kapanis_komutu(base, "Dijitalpark", "draft.xlsx",
                                            onayla=True, destek=True)
        self.assertIn("--onayla", approved)
        self.assertIn("--include-destek", approved)

    def test_kapanis_does_not_write_without_commit(self):
        with self.assertRaises(SystemExit) as raised:
            dgs_park._kapanis(None, dgs_park.PARKS["Dijitalpark"], [], "Eylül", "Eylül 2026")
        self.assertEqual(raised.exception.code, 2)

    def test_retry_preserves_draft_only_and_arge_scope(self):
        people = {
            "ORNEK ARGE": SimpleNamespace(lokasyon="Dijitalpark", arge=True),
            "ORNEK DESTEK": SimpleNamespace(lokasyon="Dijitalpark", arge=False),
        }
        data = SimpleNamespace(read_excel=lambda *_: people, _fold_tr=str.casefold,
                               is_ar_ge=lambda p: p.arge)
        def result(stdout):
            return SimpleNamespace(stdout=stdout, stderr="", returncode=0)
        calls = [result('<<<EKSIK>>>{"kisiler":["ORNEK ARGE","ORNEK DESTEK","ORNEK BILINMEYEN"]}\n'),
                 result("Giriş başarılı\n"),
                 result('<<<EKSIK>>>{"kisiler":["ORNEK BILINMEYEN"]}\n')]
        original_cwd = os.getcwd()
        with tempfile.TemporaryDirectory() as tmp:
            try:
                os.chdir(tmp)
                output = io.StringIO()
                with patch("izin_frozen.worker_cmd", return_value=["worker", "dgs"]), \
                     patch("subprocess.run", side_effect=calls) as run, \
                     contextlib.redirect_stdout(output):
                    dgs_park._kapanis(data, dgs_park.PARKS["Dijitalpark"],
                                      ["--excel", "draft.xlsx", "--commit"], "Eylül", "Eylül 2026")
                self.assertEqual(run.call_count, 3)
                retry = run.call_args_list[1].args[0]
                self.assertIn("--commit", retry)
                self.assertNotIn("--onayla", retry)
                self.assertNotIn("--include-destek", retry)
                names = Path(tmp, "dgs_kapanis_retry_Dijitalpark_Eylül.txt").read_text()
                self.assertEqual(names, "ORNEK ARGE\n")
                self.assertIn("ORNEK BILINMEYEN", output.getvalue())
                self.assertIn("MANUEL", output.getvalue())
            finally:
                os.chdir(original_cwd)


if __name__ == "__main__":
    unittest.main()
