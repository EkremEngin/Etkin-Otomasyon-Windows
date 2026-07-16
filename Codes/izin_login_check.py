#!/usr/bin/env python3
"""GUI login/park PROBE (salt-okur) — açık Chrome'daki (CDP) Arge Portal sekmesinin park'ını ve
login durumunu JSON olarak basar. `izin_gui.py` "Otomasyonu Başlat"tan ÖNCE çağırır: kullanıcının
login olup olmadığını (dashboard + Cloudflare) anlar. Portala HİÇBİR ŞEY yazmaz.

Çıktı (tek satır JSON): {"ok": bool, "park": "YTP"|null, "portal": url|null, "reason": "..."}
Kullanım: python3 izin_login_check.py [cdp_url]
"""
import os
import sys
import json

CDP = os.environ.get("DGS_CDP", "http://localhost:9222")


def probe(cdp: str) -> dict:
    out = {"ok": False, "park": None, "portal": None, "reason": ""}
    try:
        from playwright.sync_api import sync_playwright
    except Exception as e:  # noqa
        out["reason"] = f"playwright yüklü değil ({e})."
        return out
    import izin_otomasyon  # park_from_url (portalsız import; playwright'ı kendi çağırınca yükler)

    with sync_playwright() as pw:
        try:
            browser = pw.chromium.connect_over_cdp(cdp)
        except Exception as e:  # noqa
            out["reason"] = (f"Chrome'a bağlanılamadı ({cdp}). Debug-portlu Chrome açık mı? "
                             f"'Chrome'u Aç' butonuna bas. ({str(e)[:70]})")
            return out
        ctx = browser.contexts[0] if browser.contexts else None
        pages = list(ctx.pages) if ctx else []
        page = next((p for p in pages if "argeportal" in (p.url or "")), None)
        if page is None:
            out["reason"] = "Açık Arge Portal sekmesi yok. 'Chrome'u Aç' ile parkın portalını aç."
            return out
        park = izin_otomasyon.park_from_url(page.url)
        if park is None:
            out["reason"] = f"Açık sekme bilinen bir Arge Portal değil: {(page.url or '')[:60]}"
            return out
        out["park"] = park.code
        out["portal"] = park.portal_url
        # --- login kontrolü (izin_poc.assert_logged_in mantığı; exception yerine bool) ---
        try:
            title = page.title() or ""
        except Exception:
            title = ""
        body = ""
        try:
            body = page.inner_text("body")[:600]
        except Exception:
            pass
        if ("Güvenlik doğrulaması" in body or "Just a moment" in title or "Bir dakika" in title
                or "Gerçek kişi" in body or "GİRİŞ" in title.upper()):
            out["reason"] = "Cloudflare/giriş ekranı görünüyor — elle geç, dashboard açılınca tekrar dene."
            return out
        try:
            has_personel = page.locator("text=PERSONEL").count() > 0
        except Exception:
            has_personel = False
        if not has_personel:
            out["reason"] = "Dashboard görülmedi (PERSONEL menüsü yok). Portala login ol, sonra tekrar dene."
            return out
        out["ok"] = True
        out["reason"] = f"Login OK — {park.code} ({park.ad})"
        return out


def main():
    cdp = sys.argv[1] if len(sys.argv) > 1 else CDP
    try:
        res = probe(cdp)
    except Exception as e:  # noqa
        res = {"ok": False, "park": None, "portal": None, "reason": f"probe hata: {str(e)[:120]}"}
    print(json.dumps(res, ensure_ascii=False))


if __name__ == "__main__":
    main()
