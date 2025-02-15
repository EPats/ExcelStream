import tkinter as tk
import traceback

import pythoncom
import pywintypes
import win32com.client
import win32gui
from win32com.client import CDispatch


class ExcelFormulaOverlay:
    def __init__(self) -> None:
        # Excel Instance
        self.excel: CDispatch = win32com.client.Dispatch('Excel.Application')

        # Create the main window
        self.root: tk.Tk = tk.Tk()
        self.root.title('Excel Formula Live Tracker')
        self.root.attributes('-topmost', True)
        self.root.attributes('-alpha', 0.9)

        # Handle to Window IDs
        self.excel_hwnd: int | None = None
        self.formula_bar_hwnd: int | None = None
        self.formula_edit_hwnd: int | None = None

        # Reference Values and Locations
        self.last_address: str | None = None
        self.current_formula: dict[str, str] | None = None
        self.last_cell: str | None = None

        # Logic/Timers
        self.edit_mode: bool = False
        self.editing_cell: str | None = None
        self.pending_check: bool = False
        self.check_timer: str | None = None

        # Labels
        self.formula_label: tk.Label = tk.Label(
            self.root,
            text='Waiting...',
            wraplength=300,
            justify='left',
            font=('Comfortaa', 10)
        )
        self.formula_label.pack(padx=20, pady=20)

        self.find_excel_windows()
        self.update_formula()

    def find_excel_windows(self) -> None:
        def callback(hwnd: int, ctx: None) -> bool:
            try:
                try:
                    class_name: str = win32gui.GetClassName(hwnd)
                except win32gui.error as e:
                    print(f'Win32 error getting class name: {e}')
                    return True

                if class_name == 'EXCEL7':
                    self.excel_hwnd = hwnd

                    def find_formula_controls(child_hwnd: int, _: None) -> bool:
                        try:
                            class_name: str = win32gui.GetClassName(child_hwnd)
                            if class_name == 'EXCEL6.0':
                                self.formula_bar_hwnd = child_hwnd
                            elif class_name == 'EXCEL2':
                                self.formula_edit_hwnd = child_hwnd
                        except win32gui.error as e:
                            print(f'Win32 error accessing child window: {e}')
                        return True

                    try:
                        win32gui.EnumChildWindows(hwnd, find_formula_controls, None)
                    except win32gui.error as e:
                        print(f'Win32 error enumerating child windows: {e}')

            except (win32gui.error, pywintypes.error) as e:
                print(f'Windows API error in callback: {e}')
            return True

        try:
            win32gui.EnumWindows(callback, None)
        except win32gui.error as e:
            print(f'Win32 error enumerating windows: {e}')
            self.excel_hwnd = None

    def show_error_message(self, message: str) -> None:
        self.current_formula = None
        self.formula_label.config(text=f'Error: {message}')

    def get_cell_address(self, cell: CDispatch) -> str:
        return cell.Worksheet.Name + ' - ' + cell.Address.replace('$', '')

    def get_cell_formula(self, cell: CDispatch) -> str:
        return cell.Formula

    def get_cell_details(self, cell: CDispatch) -> dict[str: str]:
        return {
            'address': self.get_cell_address(cell),
            'formula': self.get_cell_formula(cell)
        }

    def update_current_formula(self, cell: CDispatch) -> None:
        self.current_formula = self.get_cell_details(cell)
        self.update_display()

    def check_cell_and_update(self, cell_ref: str | None) -> bool:
        if not cell_ref:
            return False

        cell: CDispatch = self.excel.Range(cell_ref)
        formula: str = self.get_cell_formula(cell)
        if formula.startswith('='):
            self.update_current_formula(cell)
            return True
        elif cell.HasSpill:
            parent_cell: CDispatch = cell.SpillParent
            self.update_current_formula(parent_cell)
            return True
        return False

    def get_active_cell_address(self) -> str | None:
        if self.excel is None:
            self.show_error_message('Excel not found.')
            return None
        elif self.excel.ActiveCell is None:
            self.show_error_message('Excel Workbook not found.')
            return None
        else:
            try:
                return self.excel.ActiveCell.Address
            except Exception as e:
                print(e)
                return None

    def check_for_new_formula(self) -> None:
        self.check_cell_and_update(self.editing_cell)
        self.pending_check = False
        self.check_timer = None

    def schedule_formula_check(self) -> None:
        # Schedule a check for new formula after a brief delay
        if not self.pending_check:
            self.pending_check = True
            if self.check_timer:
                self.root.after_cancel(self.check_timer)
            self.check_timer = self.root.after(500, self.check_for_new_formula)

    def update_display(self) -> None:
        display_text = f'{self.current_formula['address']}:\n{self.current_formula['formula']}'
        self.formula_label.config(text=display_text)

    def update_formula(self) -> None:
        # Check if we need to find Excel windows
        if not self.excel_hwnd or not self.formula_edit_hwnd:
            self.find_excel_windows()

        try:
            current_cell: CDispatch = self.excel.ActiveCell
            current_address: str = current_cell.Address
        except (pythoncom.com_error, AttributeError):
            # Excel not responding or no active cell
            # Most likely an open dialogue
            # We don't care, just wait until it's ready!
            self.root.after(50, self.update_formula)
            return

        previous_edit_mode: bool = self.edit_mode
        try:
            _ = current_cell.Formula
            self.edit_mode = False
        except (pythoncom.com_error, AttributeError):
            self.edit_mode = True

        if self.edit_mode and not previous_edit_mode:
            # Just entered edit mode
            self.editing_cell = current_address
        elif not self.edit_mode and previous_edit_mode:
            # Just exited edit mode
            self.schedule_formula_check()

        # Check formula of previous cell if cell changed and not in edit mode
        if not self.edit_mode and current_address != self.last_cell:
            if self.last_cell and current_address:
                self.check_cell_and_update(self.last_cell)
            self.last_cell = current_address

        # Schedule next update
        self.root.after(50, self.update_formula)

    def safe_quit(self) -> None:
        try:
            if self.check_timer:
                self.root.after_cancel(self.check_timer)
            self.excel = None
        finally:
            pass
        self.root.quit()
        self.root.destroy()

    def run(self) -> None:
        try:
            self.root.mainloop()
        except Exception as e:
            print(f'Error in mainloop: {e}')


if __name__ == '__main__':
    app = ExcelFormulaOverlay()
    app.run()