"""Klasör seçiminden yükleyiciye kadar belge kaynağının korunması."""
import contextlib
import io
import json
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import Mock, patch
from openpyxl import Workbook

from izin_data_v2 import IzinPersonV2, PARKS
from izin_gui import IzinGUI
from izin_otomasyon import run_approval_pdfli, run_plan


class IzinBelgeAkisiTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.park = PARKS["İYTE"]

    def test_folder_picker_plan_and_run_use_same_source_in_shared_park(self):
        gui = IzinGUI.__new__(IzinGUI)
        gui.root = gui.belge_lbl = Mock()
        gui._cur_park = Mock(return_value=self.park)
        gui._modal_oncesi = gui._klasoru_hatirla = gui._set_status = Mock()
        gui._baslangic_klasoru = Mock(return_value=str(self.root))
        gui.excel_path = Mock()
        gui.excel_path.get.return_value = "izin.xlsx"
        gui.commit = gui.onayla = Mock()
        gui.commit.get.return_value = True
        gui.belge_path = Mock()
        gui.belge_path.set.side_effect = lambda path: setattr(gui.belge_path.get, "return_value", path)
        with patch("izin_gui.filedialog.askdirectory", return_value=str(self.root)) as folder_picker:
            gui._pick_belge()
        folder_picker.assert_called_once()
        for plan in (True, False):
            cmd = gui._build_cmd(plan)
            self.assertEqual(cmd[cmd.index("--kisi-belge-klasoru") + 1], str(self.root))
            self.assertNotIn("--ortak-belge", cmd)
        with patch("izin_gui.filedialog.askdirectory", return_value=""):
            gui._pick_belge()
        self.assertEqual(gui.belge_path.get(), str(self.root))
        shared = str(self.root / "Ortak.pdf")
        with patch("izin_gui.filedialog.askopenfilename", return_value=shared):
            gui._pick_belge(ortak=True)
        for plan in (True, False):
            cmd = gui._build_cmd(plan)
            self.assertEqual(cmd[cmd.index("--ortak-belge") + 1], shared)
            self.assertNotIn("--kisi-belge-klasoru", cmd)

    def test_worker_receives_only_verified_paths_for_saved_people(self):
        ayse = IzinPersonV2("10000000146", "Ayşe Işık", self.park.code)
        melis = IzinPersonV2("10000000214", "Melis Yurt", self.park.code)
        for name in ("Ayse Isik Izin.pdf", "Melis Yurt.pdf"):
            (self.root / name).touch()
        received = []
        def worker(cmd, **kwargs):
            manifest = Path(cmd[cmd.index("--belge-listesi") + 1])
            received.extend(json.loads(manifest.read_text(encoding="utf-8")))
            self.assertNotIn("--dosya", cmd)
        with patch("izin_otomasyon.os.path.exists", return_value=True), \
             patch("izin_otomasyon.subprocess.run", side_effect=worker) as run, \
             contextlib.redirect_stdout(io.StringIO()):
            run_approval_pdfli(self.park, {"ay_key": "Ekim", "donem_label": "EKİM 2026"},
                               [ayse, melis], None, "", 1, kisi_klasoru=str(self.root),
                               kaydedilenler={ayse.tc})
        run.assert_called_once()
        self.assertEqual(received, [{"tc": ayse.tc, "ad": ayse.ad,
                                     "files": [str(self.root / "Ayse Isik Izin.pdf")]}])

    def test_dijitalpark_folder_preview_is_local_and_optional(self):
        park = PARKS["DIJITALPARK"]
        xl = self.root / "izin.xlsx"
        wb = Workbook()
        wb.active.append(["T.C.", "Ad-Soyad", "Tarih", "Gün Sayısı"])
        wb.active.append(["31**", "Ayşe Işık", datetime(2026, 10, 1), 1])
        wb.save(xl)
        (self.root / "Ayse Isik Izin.pdf").touch()
        gui = IzinGUI.__new__(IzinGUI)
        gui._cur_park = Mock(return_value=park)
        gui._modal_oncesi = gui._klasoru_hatirla = gui._set_status = gui._log = Mock()
        gui.root = gui.belge_lbl = Mock()
        gui._baslangic_klasoru = Mock(return_value=str(self.root))
        gui.excel_path = Mock()
        gui.excel_path.get.return_value = str(xl)
        gui.commit = gui.onayla = Mock()
        gui.commit.get.return_value = False
        gui.belge_path = Mock()
        gui.belge_path.get.return_value = str(self.root)
        with patch("izin_gui.filedialog.askdirectory", return_value=str(self.root)) as pick:
            gui._pick_belge()
        pick.assert_called_once()
        gui._check_belge()
        gui._log.assert_any_call("  Ayşe Işık → Ayse Isik Izin.pdf\n")
        self.assertIn("--kisi-belge-klasoru", gui._build_cmd(plan=True))
        self.assertNotIn("--kisi-belge-klasoru", gui._build_cmd(plan=False))
        self.assertFalse(park.onay_dogrulandi)
        with contextlib.redirect_stdout(io.StringIO()) as out:
            run_plan(str(xl), [park], None, kisi_klasoru=str(self.root))
        self.assertIn("TASLAK · yerel denetim: kişi-başı belge 1/1", out.getvalue())
        self.assertIn("PDF: Ayse Isik Izin.pdf", out.getvalue())
        with contextlib.redirect_stdout(io.StringIO()) as out:
            run_plan(str(xl), [park], None)
        self.assertIn("yalnız taslak", out.getvalue())


if __name__ == "__main__":
    unittest.main()
