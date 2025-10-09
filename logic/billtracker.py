# logic/billtracker.py
import io, os
from datetime import date
import pandas as pd
from pypdf import PdfReader
from openpyxl import load_workbook


def _page_count_from_filestorage(fs) -> int | None:
    try:
        fs.stream.seek(0)
        reader = PdfReader(fs.stream)
        if getattr(reader, "is_encrypted", False):
            try:
                reader.decrypt("")
            except Exception:
                pass
            fs.stream.seek(0)
            reader = PdfReader(fs.stream)
        return len(reader.pages)
    except Exception:
        return None


def build_excel(file_storages, filename_col="Filename"):
    rows = []
    today = date.today().strftime("%Y-%m-%d")
    for fs in file_storages:
        name = os.path.basename(fs.filename or "file.pdf")
        pc = _page_count_from_filestorage(fs)
        rows.append(
            {
                filename_col: name,
                "Page Count": pc if isinstance(pc, int) else "ERROR",
                "Complete Y/N": "",
                "Note": "",
                "Date Received": today,
                "Date Complete": "",
            }
        )

    df = pd.DataFrame(
        rows,
        columns=[
            filename_col,
            "Page Count",
            "Complete Y/N",
            "Note",
            "Date Received",
            "Date Complete",
        ],
    )

    # Write to Excel in-memory
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    # Auto-fit columns
    buf.seek(0)
    wb = load_workbook(buf)
    ws = wb.active
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_len + 2
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out
