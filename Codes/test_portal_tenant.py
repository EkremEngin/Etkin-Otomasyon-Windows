"""Dijitalpark gerçek kayıtları yanlış şirket oturumunda başlamaz."""
import unittest
from unittest.mock import patch

import dgs_poc
import izin_poc
import portal_tenant


class FakePage:
    def __init__(self, company, title=None):
        self.company = company
        self.page_title = title if title is not None else company

    def evaluate(self, _javascript):
        return {"companies": [self.company], "title": self.page_title}

    def title(self):
        return self.page_title

    def inner_text(self, _selector):
        return "PERSONEL"

    def locator(self, _selector):
        return self

    def count(self):
        return 1


class PortalTenantTests(unittest.TestCase):
    def test_exact_legal_name_required(self):
        page = FakePage("ÖRNEK PROJE ANONİM ŞİRKETİ")
        with self.assertRaises(portal_tenant.CompanyMismatch):
            portal_tenant.assert_expected_company(page, "")
        with self.assertRaises(portal_tenant.CompanyMismatch):
            portal_tenant.assert_expected_company(page, "BAŞKA FİRMA ANONİM ŞİRKETİ")
        self.assertEqual(portal_tenant.assert_expected_company(
            page, "Ornek Proje Anonim Sirketi"), page.company)
        with self.assertRaises(portal_tenant.CompanyMismatch):
            portal_tenant.assert_expected_company(
                FakePage(page.company, "FARKLI FİRMA"), "BAŞKA FİRMA ANONİM ŞİRKETİ")
        self.assertEqual(portal_tenant.resume_suffix(page.company),
                         portal_tenant.resume_suffix("Ornek Proje Anonim Sirketi"))
        self.assertNotEqual(portal_tenant.resume_suffix(page.company),
                            portal_tenant.resume_suffix("BAŞKA FİRMA ANONİM ŞİRKETİ"))

    def test_dgs_and_izin_stop_before_write(self):
        page = FakePage("BAŞKA FİRMA ANONİM ŞİRKETİ")
        for engine in (dgs_poc, izin_poc):
            with self.subTest(engine=engine.__name__), patch.dict(
                    engine.CONFIG, {"require_company_guard": True,
                                    "expected_company": "ÖRNEK PROJE ANONİM ŞİRKETİ"}):
                with self.assertRaisesRegex(engine.CloudflareHalt, "kayıt durduruldu"):
                    engine.assert_logged_in(page)


if __name__ == "__main__":
    unittest.main()
