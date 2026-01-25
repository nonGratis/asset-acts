from typing import Optional

from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload

from .config import SHARED_DRIVE_ID, log
from .formatters import fmt_number


def upload_to_drive(drive_service, file_path: str, file_name: str) -> str:
    try:
        media = MediaFileUpload(
            file_path,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        file_metadata = {"name": file_name, "parents": [SHARED_DRIVE_ID]}

        file = drive_service.files().create(
            body=file_metadata, media_body=media, supportsAllDrives=True
        ).execute()

        return file.get("id")

    except HttpError as e:
        raise RuntimeError(f"Drive upload failed: {e}")


def upload_to_drive_safe(drive_service, docx_path: str, file_name: str, code: str, 
                          items_count: int, total_sum) -> Optional[str]:
    try:
        drive_file_id = upload_to_drive(drive_service, docx_path, f"{file_name}.docx")
        log.info(
            f'Created and uploaded "{file_name}.docx" (ID: {drive_file_id}) - '
            f'items={items_count} - sum={fmt_number(total_sum)}'
        )
        return drive_file_id
    except Exception as e:
        log.warning(f"Drive upload failed for {code}: {e}")
        log.info(
            f'Created local "{docx_path}" (upload failed) - '
            f'items={items_count} - sum={fmt_number(total_sum)}'
        )
        return None
