# -*- coding: utf-8 -*-
"""Package wrapper to satisfy xlwings Ribbon importing 'VoltAmpero'."""
from voltampero import get_controller  # re-export as needed

def main():
    """xlwings Ribbon 'Run' entrypoint: attach Excel and mark status."""
    try:
        ctrl = get_controller()
        ctrl.attach_excel()
        if getattr(ctrl, 'control_sheet', None):
            ctrl.control_sheet.range("ExportStatus").value = "xlwings: OK"
    except Exception as e:
        try:
            ctrl.control_sheet.range("ExportStatus").value = f"Error: {e}"
        except Exception:
            print(e)
