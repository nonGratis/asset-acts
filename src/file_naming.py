from datetime import datetime

from .config import FILE_NAME_PATTERN


def generate_file_name(dept_code: str) -> str:
    date_str = datetime.now().strftime("%Y %m %d")
    return FILE_NAME_PATTERN.format(date=date_str, deptname=dept_code)
