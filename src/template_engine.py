from decimal import Decimal
from typing import Dict, Any, List

from docxtpl import DocxTemplate
from num2words import num2words

from .formatters import fmt_number, money_to_words
from .config import log


def build_mapping_for_owner(data: Dict[str, Any], dept: Dict[str, str]) -> Dict[str, str]:
    tot_qty = int(data.get("tot_qty", 0))
    tot_sum = data.get("tot_sum", Decimal("0.00"))

    receiver_position = dept.get("receiver_position", "")
    receiver_name = dept.get("receiver_normalized", "")
    
    if not receiver_position:
        log.warning(f"Department '{dept.get('code', '')}' has empty receiver position")
    if not receiver_name:
        log.warning(f"Department '{dept.get('code', '')}' has empty receiver name")
    
    mapping = {
        "TotalQuantityWords": num2words(tot_qty, lang="uk"),
        "TotalQuantityNumeric": str(tot_qty),
        "TotalSumNumeric": fmt_number(tot_sum),
        "TotalSumWords": money_to_words(tot_sum, lang="uk"),
        "SecondDirectorPosition": dept.get("position", ""),
        "SecondDirectorName": dept.get("normalized", ""),
        "ReceiverPosition": receiver_position,
        "ReceiverName": receiver_name,
        "Val": fmt_number(tot_sum),
    }
    return mapping


def prepare_items_for_template(items: List[Dict[str, Any]]) -> List[Dict[str, str]]:
    formatted_items = []
    for item in items:
        formatted_items.append({
            "name": str(item.get("name", "")),
            "inventory": str(item.get("inventory", "")),
            "unit": str(item.get("unit", "")),
            "qty": str(int(item.get("qty", 0))),
            "unit_price": fmt_number(item.get("unit_price", Decimal("0.00"))) if item.get("unit_price") is not None else "",
            "sum": fmt_number(item.get("sum", Decimal("0.00"))) if item.get("sum") is not None else "",
            "note": str(item.get("note", "")),
        })
    return formatted_items


def render_document(template_path: str, context: Dict[str, Any], output_path: str) -> None:
    doc = DocxTemplate(template_path)
    doc.render(context)
    doc.save(output_path)
