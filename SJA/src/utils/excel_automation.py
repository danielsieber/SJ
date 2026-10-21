"""
Excel automation via win32com for managing SchanzenID entries.
"""
import os


def _open_listen_workbook(listen_path):
    """Returns (excel, workbook), reusing an open Excel instance/workbook."""
    import win32com.client

    listen_path = os.path.abspath(listen_path)
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        excel = win32com.client.Dispatch("Excel.Application")

    excel.Visible = True
    wb = None
    for open_wb in excel.Workbooks:
        try:
            open_path = os.path.abspath(open_wb.Path + "\\" + open_wb.Name)
            if os.path.normcase(open_path) == os.path.normcase(listen_path):
                wb = open_wb
                break
        except Exception:
            pass

    if wb is None:
        wb = excel.Workbooks.Open(listen_path)

    wb.Activate()
    return excel, wb


def _bring_excel_to_foreground(excel):
    try:
        import ctypes
        hwnd = excel.Hwnd
        ctypes.windll.user32.ShowWindow(hwnd, 9)
        ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        ctypes.windll.user32.keybd_event(0x12, 0, 2, 0)
    except Exception:
        pass


def add_schanze_to_excel(listen_path, schanzen_id, hs_val):
    """
    Automates Excel using win32com to add a new SchanzenID and HS value
    to the first free cell in column A/B of sheet 'Liste_Schanzen_ID'.
    Leaves the workbook open and does NOT save it (user will review and save).
    """
    import pythoncom

    listen_path = os.path.abspath(listen_path)
    pythoncom.CoInitialize()
    try:
        excel, wb = _open_listen_workbook(listen_path)
        sheet = wb.Sheets("Liste_Schanzen_ID")
        sheet.Activate()
        
        # Find first empty row in Column A starting from row 11
        row = 11
        while sheet.Cells(row, 1).Value is not None and str(sheet.Cells(row, 1).Value).strip() != "":
            row += 1
            
        # Write values
        sheet.Cells(row, 1).Value = schanzen_id
        sheet.Cells(row, 2).Value = int(hs_val) if str(hs_val).isdigit() else hs_val
        
        # Select the new cell and scroll there
        sheet.Cells(row, 1).Select()
        try:
            excel.ActiveWindow.ScrollRow = max(1, row - 10)
        except Exception:
            pass
            
        _bring_excel_to_foreground(excel)
            
    except Exception as e:
        raise e
    finally:
        pythoncom.CoUninitialize()


def add_springer_to_excel(listen_path, springer_id, full_name, nsa):
    """
    Inserts a Springer into the alphabetically sorted SpringerID table.

    The workbook is deliberately left open and unsaved for manual review.
    """
    import pythoncom

    listen_path = os.path.abspath(listen_path)
    pythoncom.CoInitialize()
    try:
        excel, wb = _open_listen_workbook(listen_path)
        sheet = wb.Sheets("Liste_Springer_ID")
        sheet.Activate()

        target_table = None
        for idx in range(1, sheet.ListObjects.Count + 1):
            table = sheet.ListObjects.Item(idx)
            try:
                name = str(table.Name).upper()
                covers_column_a = (
                    table.Range.Column <= 1
                    and table.Range.Column + table.Range.Columns.Count - 1 >= 1
                )
                if covers_column_a and ("SPRINGER" in name or target_table is None):
                    target_table = table
                    if "SPRINGER" in name:
                        break
            except Exception:
                continue

        springer_key = str(springer_id).strip().upper()
        inserted_row = None

        if target_table is not None:
            data_count = target_table.ListRows.Count
            insert_position = data_count + 1
            for position in range(1, data_count + 1):
                value = target_table.DataBodyRange.Cells(position, 1).Value
                value_key = str(value).strip().upper() if value is not None else ""
                if not value_key or springer_key < value_key:
                    insert_position = position
                    break

            new_list_row = target_table.ListRows.Add(insert_position)
            inserted_row = new_list_row.Range.Row
        else:
            # Fallback for a sheet without a structured Excel table.
            first_data_row = 11
            last_used_row = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row  # xlUp
            inserted_row = max(first_data_row, last_used_row + 1)
            for row in range(first_data_row, last_used_row + 1):
                value = sheet.Cells(row, 1).Value
                value_key = str(value).strip().upper() if value is not None else ""
                if not value_key or springer_key < value_key:
                    inserted_row = row
                    break
            sheet.Rows(inserted_row).Insert()

        sheet.Cells(inserted_row, 1).Value = springer_id
        sheet.Cells(inserted_row, 2).Value = full_name
        sheet.Cells(inserted_row, 3).Value = nsa
        sheet.Cells(inserted_row, 1).Select()
        try:
            excel.ActiveWindow.ScrollRow = max(1, inserted_row - 10)
        except Exception:
            pass

        _bring_excel_to_foreground(excel)
    except Exception as exc:
        raise exc
    finally:
        pythoncom.CoUninitialize()
