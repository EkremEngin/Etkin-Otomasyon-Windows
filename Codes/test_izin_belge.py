"""Kişi bazında izin PDF'lerinin seçilen klasörden çözümlenmesi."""
import tempfile
import unicodedata
import unittest
import ntpath
from pathlib import Path
from unittest.mock import patch

from izin_belge import ORTAK, PER_PERSON, _pdfs, person_key, readiness, source_readiness
from izin_data_v2 import IzinPersonV2, PARKS


class IzinKlasoruTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.folder = Path(tmp.name)
        self.ayse = IzinPersonV2("10000000146", "AYŞE GÜNEŞ", "İYTE")
        self.umut = IzinPersonV2("10000000214", "UMUT UZUN", "İYTE")
        self.people = [self.ayse, self.umut]

    def file(self, name):
        path = self.folder / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.touch()
        return str(path)

    def test_recursive_turkish_names_tc_and_multiple_forms_in_number_order(self):
        first = self.file(unicodedata.normalize("NFD", "Alt/2026-Ekim_Ayşe-Güneş_İzin.PDF"))
        second = self.file("Alt/Ayşe Güneş/izin (2).pdf")
        tenth = self.file("Alt/Ayse Gunes/izin (10).PDF")
        umut = self.file("Alt/10000000214/İzinler/01.pdf")
        self.file("Alt/Puantaj.xlsx")
        self.file("Alt/__MACOSX/._UMUT UZUN.pdf")
        self.file("Alt/._Ayşe Güneş.pdf")
        self.file("Alt/Ayse Gunesli.pdf")
        self.file("Alt/_eski/UMUT UZUN.pdf")

        r = source_readiness(PARKS["İYTE"], self.people, kisi_klasoru=self.folder)
        self.assertEqual(r["mod"], PER_PERSON)
        self.assertTrue(r["hazir"])
        self.assertEqual(r["matched"], {self.ayse.tc: [first, second, tenth], self.umut.tc: [umut]})
        self.assertEqual(r["extra"], [str(Path("Alt") / "Ayse Gunesli.pdf")])
        self.assertEqual(len(_pdfs(self.folder)), 5)

    def test_conflicting_folder_name_and_id_block_even_if_separate_pdfs_exist(self):
        self.file("Ayşe Güneş.pdf")
        self.file("Umut Uzun.pdf")
        self.file("Umut Uzun/10000000146.pdf")
        self.file("Ayşe Güneş/izin Umut Uzun.pdf")
        self.file("Ayşe Güneş 10000000372.pdf")
        r = readiness(self.people, self.folder)
        self.assertEqual(r["eslesen"], 2)
        self.assertFalse(r["hazir"])
        self.assertEqual(len(r["ambiguous"]), 3)

    def test_unlisted_longer_name_does_not_match_a_shorter_employee_name(self):
        self.file("Zeynep Ayşe Güneş - Yıllık İzin.pdf")
        self.file("Zeynep Ayşe Güneş/İzin.pdf")
        form = self.file("Ekİm-2026_Ayşe-Güneş_İzin Formu.pdf")
        r = readiness([self.ayse], self.folder)
        self.assertEqual(r["matched"][self.ayse.tc], [form])
        self.assertEqual(len(r["extra"]), 2)

    def test_homonyms_require_id_and_identical_mask_prefix_keeps_distinct_people(self):
        other = IzinPersonV2(self.umut.tc, self.ayse.ad, "İYTE")
        qualified = self.file("AYŞE GÜNEŞ 10000000146.pdf")
        self.file("AYŞE GÜNEŞ.pdf")
        r = readiness([self.ayse, other], self.folder)
        self.assertFalse(r["hazir"])
        self.assertEqual(r["matched"], {self.ayse.tc: [qualified]})
        self.assertEqual(r["ambiguous"], ["AYŞE GÜNEŞ.pdf"])
        masked = [IzinPersonV2("10*****", p.ad, "İYTE") for p in self.people]
        self.file("Umut Uzun.pdf")
        masked_r = readiness(masked, self.folder)
        self.assertTrue(masked_r["hazir"])
        self.assertEqual(set(masked_r["matched"]), {person_key(p.tc, p.ad) for p in masked})

    def test_nonpdf_never_uploads_and_convertible_person_form_stays_visible(self):
        self.file("Ayşe Güneş.xlsx")
        self.file("Umut Uzun.pdf")
        self.file("Umut Uzun.xlsx")
        self.file("bilgi.txt")
        r = readiness(self.people, self.folder)
        self.assertFalse(r["hazir"])
        self.assertEqual(r["missing"], [self.ayse])
        self.assertEqual(r["nonpdf"], ["Ayşe Güneş.xlsx"])
        self.assertEqual(r["matched"][self.umut.tc], [str(self.folder / "Umut Uzun.pdf")])

    def test_personnel_folders_do_not_assign_unrelated_documents(self):
        self.file("Ayşe Güneş/Diploma.pdf")
        self.file("Ayşe Güneş/scan.pdf")
        self.file("Ayşe Güneş/İzinler/Kimlik.pdf")
        self.file("Ayşe Güneş işe giriş bildirgesi.pdf")
        form = self.file("Ayşe Güneş/Yıllık İzin Belgesi.pdf")
        r = readiness([self.ayse], self.folder)
        self.assertTrue(r["hazir"])
        self.assertEqual(r["matched"][self.ayse.tc], [form])
        self.assertEqual(len(r["extra"]), 4)

    def test_windows_converted_copy_with_different_filename_case(self):
        self.file("Ekim İzin Dosyaları/Ayşe Güneş.XLSX")
        pdf = self.file("Ekim İzin Dosyaları/ayşe güneş.PDF")
        with patch("izin_belge.os.path.normcase", ntpath.normcase):
            r = readiness([self.ayse], self.folder)
        self.assertTrue(r["hazir"])
        self.assertEqual(r["matched"][self.ayse.tc], [pdf])

    def test_same_person_forms_in_separate_folders_need_distinct_display_names(self):
        self.file("Ayşe Güneş/Ekim/İzin.pdf")
        self.file("Ayşe Güneş/Kasım/I\u0307zin.pdf")
        r = readiness([self.ayse], self.folder)
        self.assertFalse(r["hazir"])
        self.assertIn("farklı adlandırın", r["hata"])

    def test_invalid_explicit_selection_never_discovers_an_alternate_folder(self):
        self.file("İYTE İzin Belgeleri/Genel.pdf")
        r = source_readiness(PARKS["İYTE"], self.people, base=self.folder, kisi_klasoru=self.folder / "yok")
        self.assertFalse(r["hazir"])
        self.assertEqual(r["mod"], PER_PERSON)
        self.assertEqual(r["missing"], self.people)
        self.assertEqual(r["folder"], str(self.folder / "yok"))

    def test_legacy_shared_pdf_is_still_shared_with_unrelated_neighbour_files(self):
        shared = self.file("Antetli.pdf")
        self.file("Ayşe Güneş.pdf")
        self.file("Liste.xlsx")
        r = source_readiness(PARKS["İYTE"], self.people, ortak_belge=shared)
        self.assertTrue(r["hazir"])
        self.assertEqual(r["mod"], ORTAK)
        self.assertEqual(r["matched"], {p.tc: [shared] for p in self.people})
        self.assertFalse(readiness(self.people, self.folder, mode=ORTAK)["hazir"])


if __name__ == "__main__":
    unittest.main()
