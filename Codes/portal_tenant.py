"""Portal oturumundaki şirketi gerçek kayıt öncesi doğrula."""
import hashlib
import re
import unicodedata


class CompanyMismatch(ValueError):
    """Beklenen şirket ile portal oturumu eşleşmiyor."""


def _name_key(value: str) -> str:
    text = unicodedata.normalize("NFKD", str(value or "").replace("ı", "i").casefold())
    text = "".join(char for char in text if not unicodedata.combining(char))
    return " ".join(re.findall(r"\w+", text))


def resume_suffix(expected: str) -> str:
    """Aynı parktaki farklı firmaların devam kayıtlarını ayıran sabit kısa kimlik."""
    key = _name_key(expected)
    if not key:
        raise CompanyMismatch("Dijitalpark devam dosyası için beklenen tam firma unvanı girilmeli.")
    return hashlib.sha256(key.encode("utf-8")).hexdigest()[:12]


def assert_expected_company(page, expected: str) -> str:
    """Portalda görünen tek firma adını beklenen tam unvanla eşleştir."""
    if not _name_key(expected):
        raise CompanyMismatch("Dijitalpark gerçek kayıt için beklenen tam firma unvanı girilmeli.")
    state = page.evaluate("""() => ({
      companies: [...document.querySelectorAll('span.message.firm')]
        .filter(el => el.getClientRects().length > 0)
        .map(el => (el.textContent || '').trim()).filter(Boolean)
    })""")
    companies = state.get("companies", []) if isinstance(state, dict) else []
    if len(companies) != 1:
        raise CompanyMismatch("Portalda tek bir görünür firma unvanı doğrulanamadı; kayıt durduruldu.")
    actual = companies[0]
    if _name_key(actual) != _name_key(expected):
        raise CompanyMismatch(f"Portalda açık firma '{actual}', beklenen firma '{expected}' değil; kayıt durduruldu.")
    return actual
