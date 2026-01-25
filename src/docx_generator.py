import os
from typing import Dict, Any

from .config import OUTPUT_LOCAL_DIR_DOC, TEMPLATE_PATH
from .template_engine import (
    build_mapping_for_owner,
    prepare_items_for_template,
    render_document,
)


def save_docx_locally(template_path: str, output_path: str, mapping: dict, items: list):
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    context = {**mapping, "items": prepare_items_for_template(items)}
    render_document(template_path, context, output_path)


def create_docx_for_owner(code: str, data: Dict[str, Any], file_name: str) -> Dict[str, Any]:
    dept = data["dept"]
    mapping = build_mapping_for_owner(data, dept)
    docx_path = os.path.join(OUTPUT_LOCAL_DIR_DOC, f"{file_name}.docx")
    
    save_docx_locally(
        template_path=TEMPLATE_PATH,
        output_path=docx_path,
        mapping=mapping,
        items=data["items"],
    )
    
    return {
        "docx_path": docx_path,
        "name": file_name,
        "items": len(data["items"]),
        "sum": data["tot_sum"],
    }
