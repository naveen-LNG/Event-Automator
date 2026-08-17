"""Formula-safe workbook loading for every tool that writes a design .xlsx.

The problem this exists to prevent
----------------------------------
openpyxl cannot round-trip a formula AND its cached result. Loading a
workbook normally gives you the formula text; on save openpyxl writes the
formula back but drops the cached value Excel had stored alongside it.

coder_convert reads cached values, not formulas. So a plain
load_workbook(p) -> edit -> save(p) silently blanks every formula cell in
the file, and conversion then fails with a check-ref error pointing at
whichever row happens to hold the first formula -- a row that usually has
nothing to do with the event being edited.

There is a second, quieter failure in the same family: openpyxl's
insert_rows() moves formula cells without rewriting their references, so a
formula that read its own row ends up reading a row N above itself.

Both go away if formulas are baked down to their values at LOAD time,
before any edits or insertions. Coordinates still line up 1:1 between the
formula view and the value view at that point, and nothing downstream ever
sees a formula.
"""

import os

import openpyxl


def load_workbook_frozen(path, **kwargs):
    """Load a workbook with every formula replaced by its cached value.

    Use this instead of openpyxl.load_workbook() anywhere the workbook is
    going to be saved. Read-only callers should keep using
    load_workbook(..., data_only=True).

    A formula whose cached value is None is left as a formula rather than
    being overwritten with None -- that case is normally a formula which
    legitimately evaluates to blank (e.g. =F15&G15 with both cells empty),
    and blanking it would destroy real data.

    If EVERY formula in the workbook is missing its cached value, the cache
    was stripped wholesale by an earlier tool run (openpyxl clears the whole
    workbook's cache at once, so a partially-cached file cannot be a stripped
    one). That file cannot be saved safely, so it raises instead.
    """
    kwargs.pop("data_only", None)
    wb = openpyxl.load_workbook(path, **kwargs)
    wb_vals = openpyxl.load_workbook(path, data_only=True, **kwargs)

    total = 0
    stale = []
    for name in wb.sheetnames:
        if name not in wb_vals.sheetnames:
            continue
        ws, ws_vals = wb[name], wb_vals[name]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    total += 1
                    cached = ws_vals[cell.coordinate].value
                    if cached is None:
                        stale.append(f"{name}!{cell.coordinate}")
                    else:
                        cell.value = cached
    wb_vals.close()

    if total and len(stale) == total:
        raise ValueError(
            f"{os.path.basename(path)}: all {total} formula cells are missing their "
            f"cached values (e.g. {', '.join(stale[:5])}). This file was last written "
            f"by a tool that stripped formula results, so saving it again would ship "
            f"blank cells to coder_convert. Fix: revert the file, or open it in Excel "
            f"and save once to recalculate, then re-run."
        )
    return wb
