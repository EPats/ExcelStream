import tkinter as tk
import pythoncom
import win32com.client
import win32gui
import win32process
import win32api
import logging
import gc
from win32com.client import CDispatch

# Constants
WINDOW_WIDTH = 400
WINDOW_HEIGHT = 200
WINDOW_OPACITY = 0.9
TEXT_WRAP_WIDTH = 380
FORMULA_FONT = ('Consolas', 10)
STATUS_FONT = ('Consolas', 8)
UPDATE_INTERVAL_MS = 50
RETRY_INTERVAL_MS = 100
FORMULA_CHECK_DELAY_MS = 500

# Excel window class names
EXCEL_MAIN_CLASS = 'XLMAIN'
FORMULA_BAR_CLASSES = ['EXCEL6', 'EXCEL6.0']
FORMULA_EDIT_CLASSES = ['EXCEL2', 'EXCEL7', 'EXCEL9']


class ExcelFormulaOverlay:
    """
    A floating overlay window that displays the formula of the currently selected Excel cell.

    This class tracks active Excel instances, finds formula controls, and displays
    real-time formula information for the active cell.
    """

    def __init__(self) -> None:
        # Setup logging
        self.logger = logging.getLogger(__name__)

        # Initialize COM
        pythoncom.CoInitialize()  # type: ignore

        # Process Info
        self.excel_processes: dict[int, CDispatch] = {}
        self.excel_windows: dict[int, int] = {}
        self.excel_app: CDispatch | None = None

        # Window IDs
        self.active_excel_hwnd: int | None = None
        self.active_excel_pid: int | None = None
        self.formula_bar_hwnd: int | None = None
        self.formula_edit_hwnd: int | None = None

        # Cell tracking
        self.last_address: str | None = None
        self.current_formula: dict[str, str] | None = None
        self.last_cell: str | None = None

        # State tracking
        self.edit_mode: bool = False
        self.editing_cell: str | None = None
        self.pending_check: bool = False
        self.check_timer: str | None = None

        # UI setup
        self._setup_ui()

        # Start Excel tracking
        self.initialize_excel_instances()
        self.update_formula()

    def _setup_ui(self) -> None:
        """Set up the tkinter UI components for the overlay window."""
        # Main window
        self.root: tk.Tk = tk.Tk()
        self.root.title('Excel Formula Live Tracker')
        self.root.attributes('-topmost', True)
        self.root.attributes('-alpha', WINDOW_OPACITY)
        self.root.geometry(f'{WINDOW_WIDTH}x{WINDOW_HEIGHT}')

        # Formula display
        self.formula_label: tk.Label = tk.Label(
            self.root,
            text='Waiting for Excel...',
            wraplength=TEXT_WRAP_WIDTH,
            justify='left',
            font=FORMULA_FONT
        )
        self.formula_label.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Status display
        self.instance_label: tk.Label = tk.Label(
            self.root,
            text='Searching for Excel instances...',
            font=STATUS_FONT
        )
        self.instance_label.pack(pady=5, fill=tk.X)

    def initialize_excel_instances(self) -> None:
        """
        Find all running Excel instances and establish COM connections.
        Preserves the previously active Excel instance if possible.
        """
        active_pid = self.active_excel_pid
        active_hwnd = self.active_excel_hwnd

        self.excel_processes = {}
        self.excel_windows = {}
        excel_pids = self._find_excel_windows()
        excel_instances_count = self._connect_to_excel_processes(excel_pids)
        self._update_instance_info(excel_instances_count)
        self._restore_active_excel(active_pid, active_hwnd)

    def _find_excel_windows(self) -> set:
        """
        Find all Excel windows in the system.

        Returns:
            set: A set of process IDs for Excel instances.
        """
        excel_pids = set()

        def enum_windows_callback(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True

            try:
                class_name = win32gui.GetClassName(hwnd)
                if class_name == EXCEL_MAIN_CLASS:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    if pid:
                        self.excel_windows[hwnd] = pid
                        excel_pids.add(pid)
            except win32api.error as e:  # type: ignore[unresolved-import]
                self.logger.error(f'Error in window enumeration: {e}')
            return True

        win32gui.EnumWindows(enum_windows_callback, None)
        return excel_pids

    def _connect_to_excel_processes(self, excel_pids: set) -> int:
        """
        Connect to each Excel process using COM.

        Returns:
            int: Number of Excel instances successfully connected
        """
        excel_instances_count = 0

        for pid in excel_pids:
            excel_app = self._try_connect_to_excel_process(pid)
            if excel_app:
                self.excel_processes[pid] = excel_app
                excel_instances_count += 1

        return excel_instances_count

    def _try_connect_to_excel_process(self, pid: int) -> CDispatch | None:
        """
        Try to connect to a specific Excel process.

        Returns:
            Excel application COM object or None if connection failed
        """
        try:
            excel_app = win32com.client.Dispatch('Excel.Application')
            if not self._verify_excel_process(excel_app, pid):
                return None
            return excel_app
        except pythoncom.com_error as e:  # type: ignore[unresolved-import]
            self.logger.error(f'Error creating Excel dispatch: {e}')
            return None

    @staticmethod
    def _verify_excel_process(excel_app: CDispatch, target_pid: int) -> bool:
        """
        Verify that an Excel application object belongs to a specific process.

        Returns:
            bool: True if the Excel app belongs to the target process
        """
        try:
            if not hasattr(excel_app, 'Hwnd'):
                return False

            excel_hwnd = excel_app.Hwnd
            _, excel_pid = win32process.GetWindowThreadProcessId(excel_hwnd)
            return excel_pid == target_pid
        except (AttributeError, pythoncom.com_error) as e:  # type: ignore[unresolved-import]
            logging.error(f'Error verifying Excel process: {e}')
            return False

    def _restore_active_excel(self, active_pid: int | None, active_hwnd: int | None) -> None:
        """
        Attempt to restore the previously active Excel instance.
        Falls back to the first available instance if the previous one is not found.
        """

        # Try to restore from previous process ID
        if active_pid in self.excel_processes:
            self.active_excel_pid = active_pid
            self.excel_app = self.excel_processes[active_pid]
            return

        # Try to restore from previous window handle
        if active_hwnd in self.excel_windows:
            new_pid = self.excel_windows[active_hwnd]
            self.active_excel_pid = new_pid
            self.active_excel_hwnd = active_hwnd
            self.excel_app = self.excel_processes[new_pid]
            return

        # Use first available process
        if self.excel_processes:
            self.active_excel_pid = next(iter(self.excel_processes))
            self.excel_app = self.excel_processes[self.active_excel_pid]

            # Find corresponding window
            for hwnd, pid in self.excel_windows.items():
                if pid == self.active_excel_pid:
                    self.active_excel_hwnd = hwnd
                    break
            return

        # No Excel processes found
        self.active_excel_pid = None
        self.active_excel_hwnd = None
        self.excel_app = None

    def _update_instance_info(self, count: int) -> None:
        messages = {
            0: 'No Excel instances detected',
            1: '1 Excel instance connected'
        }
        self.instance_label.config(
            text=messages.get(count, f'{count} Excel instances connected')
        )

    def find_formula_controls(self, excel_hwnd: int) -> None:
        self.formula_bar_hwnd = None
        self.formula_edit_hwnd = None

        self._search_for_formula_controls(excel_hwnd)

        if not self.formula_bar_hwnd:
            self._search_in_child_containers(excel_hwnd)

        self.logger.debug(f'Formula controls: Bar={self.formula_bar_hwnd}, Edit={self.formula_edit_hwnd}')

    def _search_for_formula_controls(self, parent_hwnd: int) -> None:
        def find_controls(child_hwnd: int, _: None) -> bool:
            try:
                class_name = win32gui.GetClassName(child_hwnd)
                if class_name in FORMULA_BAR_CLASSES:
                    self.formula_bar_hwnd = child_hwnd
                    self._find_formula_edit_control(child_hwnd)
            except win32gui.error as e:  # type: ignore[unresolved-import]
                self.logger.error(f'Win32 error: {e}')
            return True

        try:
            win32gui.EnumChildWindows(parent_hwnd, find_controls, None)
        except win32api.error as e:  # type: ignore[unresolved-import]
            self.logger.error(f'Error searching for formula controls: {e}')

    def _find_formula_edit_control(self, formula_bar_hwnd: int) -> None:
        def find_edit(edit_hwnd: int, _: None) -> bool:
            try:
                edit_class = win32gui.GetClassName(edit_hwnd)
                if edit_class in FORMULA_EDIT_CLASSES:
                    self.formula_edit_hwnd = edit_hwnd
                    return False
            except win32api.error:  # type: ignore[unresolved-import]
                pass
            return True

        try:
            win32gui.EnumChildWindows(formula_bar_hwnd, find_edit, None)
        except win32api.error as e:  # type: ignore[unresolved-import]
            self.logger.error(f'Error finding formula edit control: {e}')

    def _search_in_child_containers(self, excel_hwnd: int) -> None:
        def find_container(container_hwnd: int, _: None) -> bool:
            if self.formula_bar_hwnd:
                return False  # Stop if already found
            self._search_for_formula_controls(container_hwnd)
            return not self.formula_bar_hwnd

        try:
            win32gui.EnumChildWindows(excel_hwnd, find_container, None)
        except win32api.error as e:  # type: ignore[unresolved-import]
            self.logger.error(f'Error searching child containers: {e}')

    def get_active_excel_window(self) -> tuple[int | None, int | None]:
        """
        Get the currently active Excel window and process ID.

        Returns:
            tuple: (window_handle, process_id) or (None, None) if not found
        """
        # Try foreground window first
        result = self._check_foreground_window()
        if result[0]:
            return result

        # Try last active window
        if self.active_excel_hwnd and self.active_excel_hwnd in self.excel_windows:
            return self.active_excel_hwnd, self.excel_windows[self.active_excel_hwnd]

        # Try any visible Excel window
        for hwnd, pid in self.excel_windows.items():
            if self._is_window_visible(hwnd):
                return hwnd, pid

        return None, None

    def _check_foreground_window(self) -> tuple[int | None, int | None]:
        try:
            foreground_hwnd = win32gui.GetForegroundWindow()
            current_hwnd = foreground_hwnd

            while current_hwnd:
                # Known Excel window?
                if current_hwnd in self.excel_windows:
                    return current_hwnd, self.excel_windows[current_hwnd]

                # New Excel window?
                if self._is_excel_window(current_hwnd):
                    _, pid = win32process.GetWindowThreadProcessId(current_hwnd)
                    self.excel_windows[current_hwnd] = pid
                    return current_hwnd, pid

                # Move to parent
                current_hwnd = self._get_parent_window(current_hwnd)
                if not current_hwnd:
                    break
        except win32api.error as e:  # type: ignore[unresolved-import]
            self.logger.error(f'Error checking foreground window: {e}')

        return None, None

    @staticmethod
    def _is_excel_window(hwnd: int) -> bool:
        try:
            return win32gui.GetClassName(hwnd) == EXCEL_MAIN_CLASS
        except win32api.error:  # type: ignore[unresolved-import]
            return False

    @staticmethod
    def _get_parent_window(hwnd: int) -> int | None:

        try:
            parent = win32gui.GetParent(hwnd)
            if parent == 0 or parent == hwnd:
                return None
            return parent
        except win32api.error:  # type: ignore[unresolved-import]
            return None

    @staticmethod
    def _is_window_visible(hwnd: int) -> bool:
        try:
            return win32gui.IsWindowVisible(hwnd)
        except win32api.error:  # type: ignore[unresolved-import]
            return False

    def update_active_excel(self) -> None:
        """
        Update which Excel instance is currently being tracked.
        Switches to a different Excel process if needed.
        """
        # Update which Excel we're tracking
        active_hwnd, active_pid = self.get_active_excel_window()

        if not active_hwnd or not active_pid:
            if not self.excel_processes or not self.excel_app:
                self.logger.debug('No Excel processes, reconnecting...')
                self.initialize_excel_instances()
            return

        if active_pid != self.active_excel_pid:
            self._switch_excel_process(active_hwnd, active_pid)
            return

        if active_hwnd != self.active_excel_hwnd:
            self.active_excel_hwnd = active_hwnd
            self.find_formula_controls(active_hwnd)

    def _switch_excel_process(self, hwnd: int, pid: int) -> None:
        # Process already connected
        if pid in self.excel_processes:
            self._activate_excel_process(hwnd, pid)
            return

        self.logger.info(f'New Excel process {pid}, reconnecting...')
        self.initialize_excel_instances()

    def _activate_excel_process(self, hwnd: int, pid: int) -> None:
        self.active_excel_pid = pid
        self.excel_app = self.excel_processes[pid]
        self.active_excel_hwnd = hwnd

        # Reset tracking state
        self.last_address = None
        self.last_cell = None
        self.edit_mode = False
        self.editing_cell = None
        self.current_formula = None

        window_title = win32gui.GetWindowText(hwnd)
        self.logger.info(f'Switched to Excel: {window_title} (pid: {pid})')

    def show_error_message(self, message: str) -> None:
        self.current_formula = None
        self.formula_label.config(text=f'Error: {message}')

    @staticmethod
    def get_cell_address(cell: CDispatch) -> str:
        return cell.Worksheet.Name + ' - ' + cell.Address.replace('$', '')

    @staticmethod
    def get_cell_formula(cell: CDispatch) -> str:
        return cell.Formula

    def get_cell_details(self, cell: CDispatch) -> dict[str, str]:
        return {
            'address': self.get_cell_address(cell),
            'formula': self.get_cell_formula(cell)
        }

    def update_current_formula(self, cell: CDispatch) -> None:
        self.current_formula = self.get_cell_details(cell)
        self.update_display()

    def check_cell_and_update(self, cell_ref: str | None) -> bool:
        """
        Check if a cell contains a formula and update the display.

        Returns:
            bool: True if a formula was found and displayed
        """
        if not cell_ref or not self.excel_app:
            return False

        active_workbook = self._get_active_workbook()
        if not active_workbook:
            return False

        active_sheet = self._get_active_sheet(active_workbook)
        if not active_sheet:
            return False

        cell = self._get_cell(active_sheet, cell_ref)
        if not cell:
            return False

        return self._check_for_formula_or_spill(cell)

    def _get_cell(self, sheet: CDispatch, cell_ref: str) -> CDispatch | None:
        """
        Get a cell object from a sheet.

        Returns:
            Cell COM object or None if not found
        """
        try:
            return sheet.Range(cell_ref)
        except pythoncom.com_error:  # type: ignore[unresolved-import]
            try:
                return self.excel_app.Range(cell_ref)
            except pythoncom.com_error:  # type: ignore[unresolved-import]
                self.logger.error(f'Could not find cell {cell_ref}')
                return None

    def _check_for_formula_or_spill(self, cell: CDispatch) -> bool:
        # Check for formula
        try:
            formula = self.get_cell_formula(cell)
            if formula.startswith('='):
                self.update_current_formula(cell)
                return True
        except pythoncom.com_error as e:  # type: ignore[unresolved-import]
            self.logger.error(f'Error checking formula: {e}')

        # Check for spill
        try:
            if hasattr(cell, 'HasSpill') and cell.HasSpill:
                parent_cell = cell.SpillParent
                self.update_current_formula(parent_cell)
                return True
        except pythoncom.com_error as e:  # type: ignore[unresolved-import]
            self.logger.error(f'Error checking spill: {e}')

        return False

    def _get_active_workbook(self) -> CDispatch | None:
        """
        Get the active workbook in Excel.

        Returns:
            Workbook COM object or None if not found
        """
        try:
            return self.excel_app.ActiveWorkbook
        except pythoncom.com_error:  # type: ignore[unresolved-import]
            # Try any open workbook
            try:
                if hasattr(self.excel_app, 'Workbooks') and self.excel_app.Workbooks.Count > 0:
                    return self.excel_app.Workbooks(1)
            except (AttributeError, pythoncom.com_error):  # type: ignore[unresolved-import]
                pass
        return None

    def _get_active_sheet(self, workbook: CDispatch) -> CDispatch | None:
        """
        Get the active worksheet in a workbook.

        Returns:
            Worksheet COM object or None if not found
        """
        try:
            return self.excel_app.ActiveSheet
        except pythoncom.com_error:  # type: ignore[unresolved-import]
            # Try first sheet
            try:
                if hasattr(workbook, 'Worksheets') and workbook.Worksheets.Count > 0:
                    return workbook.Worksheets(1)
            except (AttributeError, pythoncom.com_error):  # type: ignore[unresolved-import]
                pass
        return None

    def get_active_cell_address(self) -> str | None:
        if not self.excel_app:
            self.show_error_message('No active Excel instance.')
            return None

        try:
            if self.excel_app.ActiveCell:
                return self.excel_app.ActiveCell.Address
        except pythoncom.com_error as e:  # type: ignore[unresolved-import]
            self.logger.error(f'Error getting active cell: {e}')

        return None

    def check_for_new_formula(self) -> None:
        self.check_cell_and_update(self.editing_cell)
        self.pending_check = False
        self.check_timer = None

    def schedule_formula_check(self) -> None:
        if not self.pending_check:
            self.pending_check = True
            if self.check_timer:
                self.root.after_cancel(self.check_timer)
            self.check_timer = self.root.after(FORMULA_CHECK_DELAY_MS, self.check_for_new_formula)

    def update_display(self) -> None:
        if self.current_formula:
            display_text = f"{self.current_formula['address']}:\n{self.current_formula['formula']}"
            self.formula_label.config(text=display_text)

    def update_formula(self) -> None:
        """
        This is our main update loop.
        Check for changes and update as required.
        """
        self.update_active_excel()
        self._update_status_display()

        # No Excel instance? Wait and retry
        if not self.excel_app:
            self.root.after(RETRY_INTERVAL_MS, self.update_formula)
            return

        cell_info = self._get_current_cell_info()
        # No cell info? Wait and retry
        if not cell_info:
            self.root.after(RETRY_INTERVAL_MS, self.update_formula)
            return

        current_cell, current_address = cell_info

        self._handle_edit_mode(current_cell, current_address)
        if not self.edit_mode and current_address != self.last_cell:
            if self.last_cell and current_address:
                self.logger.debug(f'Cell changed: {self.last_cell} → {current_address}')
                self.check_cell_and_update(self.last_cell)
            self.last_cell = current_address

        # Schedule next update
        self.root.after(UPDATE_INTERVAL_MS, self.update_formula)

    def _get_current_cell_info(self) -> tuple[CDispatch, str] | None:
        """
        Get information about the current active cell in Excel.

        Returns:
            tuple: (cell_object, cell_address) or None if not available
        """
        try:
            # Check if Excel is accessible
            if not self.excel_app.Visible:
                self.logger.debug('Excel not visible')
                return None

            # Check for active workbook
            if not self.excel_app.ActiveWorkbook:
                self.formula_label.config(text='No active workbook')
                return None

            # Get active cell
            current_cell = self.excel_app.ActiveCell
            if not current_cell:
                return None

            return current_cell, current_cell.Address
        except (AttributeError, pythoncom.com_error) as e:  # type: ignore[unresolved-import]
            # Excel might be in a state we can't access
            self.logger.debug(f"COM error accessing Excel: {e}")
            return None

    def _handle_edit_mode(self, cell: CDispatch, address: str) -> None:

        previous_edit_mode = self.edit_mode
        self.edit_mode = not self._can_access_formula(cell)

        # Just entered edit mode
        if self.edit_mode and not previous_edit_mode:
            self.editing_cell = address
            self.logger.debug(f'Editing cell: {address}')

        # Just exited edit mode
        elif not self.edit_mode and previous_edit_mode:
            self.logger.debug(f'Finished editing: {self.editing_cell}')
            self.schedule_formula_check()

    @staticmethod
    def _can_access_formula(cell: CDispatch) -> bool:
        """
        Really hacky approach to find if we can access the formula.

        Instead of finding a way to actually see if it is accessible through a property,
        we try to access it and if we don't through an error then it's accessible!

        ... there must be a better way
        """
        try:
            _ = cell.Formula
            return True
        except pythoncom.com_error:  # type: ignore[unresolved-import]
            return False

    def _update_status_display(self) -> None:
        """
        Updates the status display to show which window we are tracking.

        This is only really useful for debugging.

        TODO: add in a debug state to turn this on or off.
        """
        try:
            if not self.active_excel_hwnd:
                count = len(self.excel_processes)
                if count > 0:
                    self.instance_label.config(text=f'{count} Excel instances, none active')
                else:
                    self.instance_label.config(text='No Excel instances detected')
                return

            window_title = win32gui.GetWindowText(self.active_excel_hwnd)
            process_count = len(self.excel_processes)

            if process_count > 1:
                self.instance_label.config(text=f'Tracking: {window_title} ({process_count} Excel instances)')
            else:
                self.instance_label.config(text=f'Tracking: {window_title}')

        except win32api.error as e:  # type: ignore[unresolved-import]
            self.logger.error(f'Error updating status: {e}')

    def safe_quit(self) -> None:
        """Safely shut down the application, releasing all resources."""
        self._cancel_timers()
        self._release_com_objects()
        self._close_window()

    def _cancel_timers(self) -> None:
        try:
            if self.check_timer:
                self.root.after_cancel(self.check_timer)
        except Exception as e:
            self.logger.error(f'Error canceling timers: {e}')

    def _release_com_objects(self) -> None:
        try:
            # Clear all references to COM objects
            self.excel_processes = {}
            self.excel_windows = {}
            self.excel_app = None

            # Force garbage collection to ensure COM objects are released promptly
            gc.collect()
            pythoncom.CoUninitialize()  # type: ignore
        except Exception as e:
            self.logger.error(f'Error releasing COM objects: {e}')

    def _close_window(self) -> None:
        try:
            self.root.quit()
            self.root.destroy()
        except Exception as e:
            self.logger.error(f'Error closing window: {e}')

    def run(self) -> None:
        self.root.protocol('WM_DELETE_WINDOW', self.safe_quit)
        try:
            self.root.mainloop()
        except Exception as e:
            self.logger.error(f'Error in mainloop: {e}')
            self.safe_quit()


if __name__ == '__main__':
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )

    app = ExcelFormulaOverlay()
    app.run()
