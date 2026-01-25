from typing import Dict, Any, List, Optional

from .config import log
from .formatters import fmt_number
from .file_naming import generate_file_name
from .docx_generator import create_docx_for_owner
from .drive_uploader import upload_to_drive_safe
from .pdf_converter import convert_to_pdf, convert_to_jpeg


def convert_to_pdf_and_jpeg(docx_path: str, code: str, doc_info: Dict[str, Any]) -> None:
    items_count = doc_info["items"]
    total_sum = doc_info["sum"]
    
    try:
        pdf_path = convert_to_pdf(docx_path)
        doc_info["pdf_path"] = pdf_path
        
        try:
            jpeg_path = convert_to_jpeg(pdf_path)
            doc_info["jpeg_path"] = jpeg_path
            log.info(
                f'Created docs "{docx_path}" + PDF + JPEG - '
                f'items={items_count} - sum={fmt_number(total_sum)}'
            )
        except Exception as jpeg_err:
            log.warning(f"JPEG conversion failed for {code}: {jpeg_err}")
            log.info(
                f'Created docs "{docx_path}" + PDF - '
                f'items={items_count} - sum={fmt_number(total_sum)}'
            )
    except Exception as e:
        log.warning(f"PDF conversion failed for {code}: {e}")
        log.info(
            f'Created doc "{docx_path}" (PDF/JPEG skipped) - '
            f'items={items_count} - sum={fmt_number(total_sum)}'
        )


def process_single_owner(code: str, data: Dict[str, Any], drive_service) -> Optional[Dict[str, Any]]:
    if not data["items"]:
        log.info(f"Owner {code} has no items; skipping.")
        return None
    
    dept = data["dept"]
    file_name = generate_file_name(dept.get("code"))
    
    try:
        doc_info = create_docx_for_owner(code, data, file_name)
        
        drive_file_id = upload_to_drive_safe(
            drive_service, 
            doc_info["docx_path"], 
            file_name, 
            code,
            doc_info["items"],
            doc_info["sum"]
        )
        
        if drive_file_id:
            doc_info["drive_file_id"] = drive_file_id
        
        convert_to_pdf_and_jpeg(doc_info["docx_path"], code, doc_info)
        
        return doc_info
        
    except Exception as e:
        log.error(f"Document creation failed for {code}: {e}")
        return None


def create_act_docs_local(per_owner: Dict[str, Any], drive_service) -> List[Dict[str, Any]]:
    created = []
    
    for code, data in per_owner.items():
        doc_info = process_single_owner(code, data, drive_service)
        if doc_info:
            created.append(doc_info)
    
    return created
