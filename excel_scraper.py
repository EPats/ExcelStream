import pythoncom
import win32com.client
import win32gui
import win32process
import win32api
import logging
from win32com.client import CDispatch

"""
Handles interaction with Excel instances via COM and Windows API.
Tracks Excel windows, processes, and formula information.
"""

# Constants
EXCEL_MAIN_CLASS = 'XLMAIN'


class ExcelScraper:

    def __init__(self) -> None:

        self.logger = logging.getLogger(__name__)

        pythoncom.CoInitialize()  # type: ignore

        # Process Info
        self.excel_processes: dict[int, CDispatch] = {}
        self.excel_windows: dict[int, int] = {}
        self.excel_app: CDispatch | None = None

        # Window IDs
        self.active_excel_hwnd: int | None = None
        self.active_excel_pid: int | None = None

        # Initialize Excel tracking
        self.initialize_excel_instances()

    def initialize_excel_instances(self) -> int:
        """
        Find all running Excel instances and establish COM connections.
        Keeps the last active Excel instance if possible.

        Returns the number of instances found.
        """
        active_pid = self.active_excel_pid
        active_hwnd = self.active_excel_hwnd

        self.excel_processes = {}
        self.excel_windows = {}
        excel_pids = self._find_excel_windows()
        excel_instances_count = self._connect_to_excel_processes(excel_pids)
        self._restore_active_excel(active_pid, active_hwnd)

        return excel_instances_count

    def _find_excel_windows(self) -> set[int]:
        """
        Find all Excel windows in the system.

        Returns a set of process IDs for found Excel instances.
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

    def _connect_to_excel_processes(self, excel_pids: set[int]) -> int:
        """
        Connect to each Excel process using COM.

        Returns the number of Excel instances successfully connected
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
        Returns the process if successful, otherwise returns None.
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

        Returns True if the Excel app belongs to the target process
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

    def get_active_excel_window(self) -> tuple[int | None, int | None]:
        """
        Get the currently active Excel window and process ID as a tuple, or None, None if not found
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

    def update_active_excel(self) -> bool:
        """
        Update which Excel instance is currently being tracked.
        Switches to a different Excel process if needed.

        Returns True if an active Excel instance is available
        """
        # Update which Excel we're tracking
        active_hwnd, active_pid = self.get_active_excel_window()

        if not active_hwnd or not active_pid:
            if not self.excel_processes or not self.excel_app:
                self.logger.debug('No Excel processes, reconnecting...')
                self.initialize_excel_instances()
            return False

        if active_pid != self.active_excel_pid:
            self._switch_excel_process(active_hwnd, active_pid)
            return True

        if active_hwnd != self.active_excel_hwnd:
            self.active_excel_hwnd = active_hwnd

        return True

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

        window_title = win32gui.GetWindowText(hwnd)
        self.logger.info(f'Switched to Excel: {window_title} (pid: {pid})')

    @staticmethod
    def get_cell_address(cell: CDispatch) -> str:
        """Get a formatted address for a cell including worksheet name"""
        return cell.Worksheet.Name + ' - ' + cell.Address.replace('$', '')

    @staticmethod
    def get_cell_formula(cell: CDispatch) -> str:
        """Get the formula for a cell"""
        return cell.Formula

    def get_cell_details(self, cell: CDispatch) -> dict[str, str]:
        """Get all relevant details about a cell"""
        return {
            'address': self.get_cell_address(cell),
            'formula': self.get_cell_formula(cell)
        }

    def check_cell_for_formula(self, cell_ref: str | None) -> dict[str, str] | None:
        """
        Check if a cell contains a formula.

        Returns formula details as a dict if found, None otherwise
        """
        if not cell_ref or not self.excel_app:
            return None

        active_workbook = self._get_active_workbook()
        if not active_workbook:
            return None

        active_sheet = self._get_active_sheet(active_workbook)
        if not active_sheet:
            return None

        cell = self._get_cell(active_sheet, cell_ref)
        if not cell:
            return None

        formula_details = self._check_for_formula_or_spill(cell)
        return formula_details

    def _get_cell(self, sheet: CDispatch, cell_ref: str) -> CDispatch | None:
        try:
            return sheet.Range(cell_ref)
        except pythoncom.com_error:  # type: ignore[unresolved-import]
            try:
                return self.excel_app.Range(cell_ref)
            except pythoncom.com_error:  # type: ignore[unresolved-import]
                self.logger.error(f'Could not find cell {cell_ref}')
                return None

    def _check_for_formula_or_spill(self, cell: CDispatch) -> dict[str, str] | None:
        # Check for formula
        try:
            formula = self.get_cell_formula(cell)
            if formula.startswith('='):
                return self.get_cell_details(cell)
        except pythoncom.com_error as e:  # type: ignore[unresolved-import]
            self.logger.error(f'Error checking formula: {e}')

        # Check for spill
        try:
            if hasattr(cell, 'HasSpill') and cell.HasSpill:
                parent_cell = cell.SpillParent
                return self.get_cell_details(parent_cell)
        except pythoncom.com_error as e:  # type: ignore[unresolved-import]
            self.logger.error(f'Error checking spill: {e}')

        return None

    def _get_active_workbook(self) -> CDispatch | None:
        """ Get the active workbook in Excel, or None if not found """
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
        """ Get the active worksheet in a workbook, or None if not found """
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

    def get_active_cell_info(self) -> tuple[CDispatch | None, str | None]:
        """ Get information about the current active cell in Excel as a tuple, or None, None if not available. """
        if not self.excel_app:
            return None, None

        try:
            # Check if Excel is accessible
            if not self.excel_app.Visible:
                self.logger.debug('Excel not visible')
                return None, None

            # Check for active workbook
            if not self.excel_app.ActiveWorkbook:
                self.logger.debug('No active workbook')
                return None, None

            # Get active cell
            current_cell = self.excel_app.ActiveCell
            if not current_cell:
                return None, None

            return current_cell, current_cell.Address
        except (AttributeError, pythoncom.com_error) as e:  # type: ignore[unresolved-import]
            # Excel might be in a state we can't access
            self.logger.debug(f"COM error accessing Excel: {e}")
            return None, None

    def check_edit_mode(self, cell: CDispatch) -> bool:
        """ Is in edit mode if we can't access the formula """
        return not self._can_access_formula(cell)

    @staticmethod
    def _can_access_formula(cell: CDispatch) -> bool:
        """
        Hacky approach to finding if we can access the formula or not.

        I don't like it. There must be a way we can do this through a property.
        But, it works, and it's all I can find for now.

        Essentially, try to access it. If it throws an error, then we can't access.
        """
        try:
            _ = cell.Formula
            return True
        except pythoncom.com_error:  # type: ignore[unresolved-import]
            return False

    def get_excel_window_title(self) -> str | None:
        """Safely get the title of the active Excel window"""
        if not self.active_excel_hwnd:
            return None

        try:
            return win32gui.GetWindowText(self.active_excel_hwnd)
        except win32api.error:  # type: ignore[unresolved-import]
            return None

    def get_excel_process_count(self) -> int:
        return len(self.excel_processes)

    def release_resources(self) -> None:
        """ Previous method force quit Excel. Let's not do that anymore... """
        try:
            # Clear all references to COM objects
            self.excel_processes = {}
            self.excel_windows = {}
            self.excel_app = None

            # Force garbage collection
            import gc
            gc.collect()
            pythoncom.CoUninitialize()  # type: ignore
        except Exception as e:
            self.logger.error(f'Error releasing COM objects: {e}')