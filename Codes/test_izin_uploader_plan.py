import json
import tempfile
import unittest
from pathlib import Path

from izin_onaya_dosyali_yildiz import load_file_plan, saved_portal_people


class UploadPlanTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.plan = self.folder / "plan.json"

    def entry(self, tc, ad, filename):
        path = self.folder / filename
        path.touch()
        return {"tc": tc, "ad": ad, "files": [str(path)]}

    def load(self, rows, done, person=""):
        self.plan.write_text(json.dumps(rows), encoding="utf-8")
        return load_file_plan(self.plan, done, person)

    def test_only_exact_done_and_scoped_people_receive_their_planned_files(self):
        one = self.entry("12500000084", "Sema Işık", "form1.pdf")
        two = self.entry("12600000084", "Sema Işıkçı", "form2.pdf")
        three = self.entry("12700000084", "Oya Kök", "form3.pdf")
        rows = [one, two, three]
        done = [(one["tc"], "Sema IŞIK"), ("", two["ad"]), ("", "Liste Dışı")]
        self.assertEqual(self.load(rows, done), {"sema isik\t1284": one, "sema isikci\t1284": two})
        self.assertEqual(self.load(rows, done, one["tc"]), {"sema isik\t1284": one})
        self.assertFalse(self.load(rows, done, "Sema"))
        self.assertFalse(self.load(rows, [(two["tc"], one["ad"])]))

    def test_distinct_names_with_same_prefix_only_mask(self):
        rows = [self.entry("24*****", "Gül Yıldız", "bir.pdf"),
                self.entry("24*****", "Ayşe Güneş", "iki.pdf")]
        done = [(r["tc"], r["ad"]) for r in rows]
        self.assertEqual(set(self.load(rows, done)), {"gul yildiz\t24", "ayse gunes\t24"})

    def test_rejects_conflicting_portal_identity_and_ambiguous_legacy_name(self):
        one = self.entry("12500000084", "Eda Yılmaz", "evrak1.pdf")
        two = self.entry("12600000084", "Eda Yılmaz", "evrak2.pdf")
        with self.assertRaises(ValueError):
            self.load([one, two], [(one["tc"], one["ad"])], one["tc"])
        two["tc"] = "19******52"
        with self.assertRaises(ValueError):
            self.load([one, two], [("", one["ad"])])
        self.assertEqual(self.load([one, two], [(two["tc"], two["ad"])]), {"eda yilmaz\t1952": two})

    def test_missing_or_shared_file_never_falls_back(self):
        row = self.entry("24*****", "Nehir Çelik", "form.pdf")
        other = dict(row, ad="İlke Gül")
        with self.assertRaises(ValueError):
            self.load([row, other], [("", row["ad"]), ("", other["ad"])])
        Path(row["files"][0]).unlink()
        with self.assertRaises(ValueError):
            self.load([row], [("", row["ad"])])

    def test_verified_draft_id_bridges_a_changed_surname_only_in_its_park_and_period(self):
        row = self.entry("12500000084", "Burcu Çimen Engin", "izin.pdf")
        done = [(row["tc"], row["ad"])]
        self.plan.write_text(json.dumps([row]), encoding="utf-8")
        record = dict(tc=row["tc"], ad=row["ad"], commit=True, durum="girildi", kaydedildi=True,
                      park="YTP", donem="EKİM 2026", portal_person={"ad": "Burcu Çimen", "id": "376"})
        journal = self.folder / "günlük.jsonl"
        journal.write_text("\n".join(json.dumps(r) for r in (
            record, dict(record, commit=False), dict(record, park="ULUTEK"),
            dict(record, donem="EKİM 2025"), dict(record, kaydedildi=False))), encoding="utf-8")
        known = saved_portal_people(journal, "YTP", "EKİM 2026")
        self.assertFalse(saved_portal_people(journal, "YTP", "KASIM 2026"))
        self.assertEqual(load_file_plan(self.plan, done, saved_people=known),
                         {"burcu cimen\t1284": dict(row, portal_id="376")})
        self.assertEqual(load_file_plan(self.plan, [("", row["ad"])], saved_people=known),
                         {"burcu cimen engin\t1284": row})
        self.assertEqual(load_file_plan(self.plan, done, saved_people={
            (row["tc"], "burcu cimen engin"): {("burcu yabancı", "376")}}),
            {"burcu cimen engin\t1284": row})
        with self.assertRaisesRegex(ValueError, "ayırt edilemeyen"):
            load_file_plan(self.plan, done + [("12600000084", "Burcu Çimen")], saved_people=known)
        known[(row["tc"], "burcu cimen engin")].add(("burcu cimen", "377"))
        with self.assertRaisesRegex(ValueError, "çelişen"):
            load_file_plan(self.plan, done, saved_people=known)


if __name__ == "__main__":
    unittest.main()
