import re
from decimal import Decimal, ROUND_HALF_UP
from num2words import num2words

from .config import THOUSAND_SEPARATOR, DECIMAL_SEPARATOR, CURRENCY_SUFFIX
from .data_utils import quantize_money


def fmt_number(val: Decimal) -> str:
    q = quantize_money(Decimal(val))
    s = f"{q:,.2f}"
    s = s.replace(",", "TEMP_THOUS").replace(".", "TEMP_DEC")
    s = s.replace("TEMP_THOUS", THOUSAND_SEPARATOR).replace("TEMP_DEC", DECIMAL_SEPARATOR)
    return s + (CURRENCY_SUFFIX or "")


def money_to_words(amount: Decimal, lang: str = "uk") -> str:
    q = quantize_money(amount)
    total_kop = int((q * 100).to_integral_value(rounding=ROUND_HALF_UP))
    hryv = total_kop // 100
    kop = total_kop % 100

    def _form_for(n: int, forms: tuple) -> str:
        """Select proper grammatical form for Ukrainian numerals."""
        n = abs(int(n))
        if n % 10 == 1 and n % 100 != 11:
            return forms[0]
        if n % 10 in (2, 3, 4) and n % 100 not in (12, 13, 14):
            return forms[1]
        return forms[2]

    if hryv == 0:
        hryv_words = "нуль"
    else:
        thousands = hryv // 1000
        rest = hryv % 1000
        parts = []
        if thousands:
            parts.append(num2words(thousands, lang=lang))
            parts.append(_form_for(thousands, ("тисяча", "тисячі", "тисяч")))
        if rest:
            parts.append(num2words(rest, lang=lang))
        hryv_words = " ".join(parts)

    hryv_words = re.sub(r'\bодин\b', 'одна', hryv_words)
    hryv_words = re.sub(r'\bдва\b', 'дві', hryv_words)

    return f"{hryv_words} грн. {kop:02d} коп."
