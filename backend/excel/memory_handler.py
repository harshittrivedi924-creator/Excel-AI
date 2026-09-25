"""In-memory ExcelHandler used by the Excel Add-in API.

The Excel Add-in task pane reads a *snapshot* of the currently-open
worksheet through Office.js (headers + data rows + origin coordinate) and
sends it to ``POST /api/excel/command``. This handler materializes that
snapshot into an openpyxl workbook so the existing Copilot pipeline
(parser -> validator -> calculator -> handler) runs untouched against it.

Cells are written at the real worksheet coordinate space (``origin``), so
every range / cell reference produced downstream (``=SUM(B2:B4)``,
``D21``, ...) maps 1:1 onto the live workbook that the Add-in will modify
with Office.js.
"""

import openpyxl

from backend.excel.handler import ExcelHandler


def is_blank(value):
    """True when a snapshot value carries no cell content.

    Office.js returns ``null`` for empty cells, but the WebView2 runtime
    hands some builds back an empty string instead. Both mean "no data",
    so they must never be materialised as cell content: a blank cell that
    looks occupied would inflate used ranges, count as data for
    overwrite checks, and produce spurious writes.
    """
    return value is None or value == ""


class MemoryExcelHandler(ExcelHandler):
    """ExcelHandler backed by an in-memory workbook populated from a
    header/data snapshot captured from the live workbook via Office.js.
    """

    MAX_SHEET_TITLE_LEN = 31
    ILLEGAL_TITLE_CHARS = "[]:*?/\\"

    def __init__(self, headers=None, rows=None, origin_row=1, origin_col=1, sheet_name="Sheet1"):
        super().__init__(filepath=None)
        self.filepath = None
        self.workbook = openpyxl.Workbook()
        self.sheet_name = self._safe_title(sheet_name)
        self.origin_row = int(origin_row or 1)
        self.origin_col = int(origin_col or 1)
        ws = self.workbook.active
        ws.title = self.sheet_name
        self._populate(ws, headers, rows, self.origin_row, self.origin_col)

    def _safe_title(self, name):
        """Return a valid openpyxl sheet title from an arbitrary name."""
        if not name:
            return "Sheet1"
        title = str(name)
        for ch in self.ILLEGAL_TITLE_CHARS:
            title = title.replace(ch, "_")
        title = title.strip() or "Sheet1"
        return title[: self.MAX_SHEET_TITLE_LEN]

    def _populate(self, ws, headers, rows, origin_row, origin_col):
        """Write the snapshot into the worksheet at real coordinates.

        ``headers`` (when present) are written on the ``origin_row`` row
        and the data ``rows`` follow immediately below it.
        """
        for i, header in enumerate(headers or []):
            ws.cell(row=origin_row, column=origin_col + i, value=header)
        data_start = origin_row + (1 if headers else 0)
        for r, row in enumerate(rows or []):
            for c, value in enumerate(row):
                if is_blank(value):
                    continue
                ws.cell(row=data_start + r, column=origin_col + c, value=value)
