import logging

from excel_scraper import ExcelScraper
from formula_overlay import create_formula_display

# Constants
UPDATE_INTERVAL_MS = 50
RETRY_INTERVAL_MS = 100
FORMULA_CHECK_DELAY_MS = 500

"""
Main application class that coordinates between the Excel scraper and the formula display.

This class handles the main update loop, managing the state and coordinating
interaction between components.
"""


class ExcelFormulaTracker:

    def __init__(self, display_type: str = 'tkinter', debug_mode: bool = False) -> None:
        self.logger = logging.getLogger(__name__)
        self.debug_mode: bool = debug_mode

        self.scraper = ExcelScraper()
        self.display = create_formula_display(display_type, debug_mode=debug_mode)

        self.current_formula: dict[str, str] | None = None
        self.last_cell: str | None = None
        self.edit_mode: bool = False
        self.editing_cell: str | None = None
        self.pending_check: bool = False
        self.check_timer: str | None = None

        self.display.set_close_handler(self.safe_quit)
        self.update_formula()

    def update_formula(self) -> None:
        """
        Main update loop. Checks Excel status and updates the display.
        """
        excel_active = self.scraper.update_active_excel()

        if self.debug_mode:
            self._update_status_display()

        # No Excel instance? Wait and retry
        if not excel_active:
            self.display.schedule_update(self.update_formula, RETRY_INTERVAL_MS)
            return

        cell_info = self.scraper.get_active_cell_info()
        # No cell info? Wait and retry
        if not cell_info[0]:
            self.display.schedule_update(self.update_formula, RETRY_INTERVAL_MS)
            return

        current_cell, current_address = cell_info

        # Handle edit mode changes
        self._handle_edit_mode(current_cell, current_address)

        # Handle cell changes (when not in edit mode)
        if not self.edit_mode and current_address != self.last_cell:
            if self.last_cell and current_address:
                self.logger.debug(f'Cell changed: {self.last_cell} → {current_address}')
                self._check_and_update_formula(self.last_cell)
            self.last_cell = current_address

        # Schedule next update
        self.display.schedule_update(self.update_formula, UPDATE_INTERVAL_MS)

    def _handle_edit_mode(self, cell, address: str) -> None:
        previous_edit_mode = self.edit_mode
        self.edit_mode = self.scraper.check_edit_mode(cell)

        # Just entered edit mode
        if self.edit_mode and not previous_edit_mode:
            self.editing_cell = address
            self.logger.debug(f'Editing cell: {address}')

        # Just exited edit mode
        elif not self.edit_mode and previous_edit_mode:
            self.logger.debug(f'Finished editing: {self.editing_cell}')
            self._schedule_formula_check()

    def _check_and_update_formula(self, cell_ref: str | None) -> bool:
        formula_data = self.scraper.check_cell_for_formula(cell_ref)
        if formula_data:
            self.current_formula = formula_data
            self.display.update_formula(formula_data)
            return True
        return False

    def _schedule_formula_check(self) -> None:
        if not self.pending_check:
            self.pending_check = True
            if self.check_timer:
                self.display.cancel_update(self.check_timer)
            self.check_timer = self.display.schedule_update(self._check_after_edit, FORMULA_CHECK_DELAY_MS)

    def _check_after_edit(self) -> None:
        self._check_and_update_formula(self.editing_cell)
        self.pending_check = False
        self.check_timer = None

    def _update_status_display(self) -> None:
        excel_count = self.scraper.get_excel_process_count()

        if excel_count == 0:
            self.display.update_status('No Excel instances detected')
            return

        window_title = self.scraper.get_excel_window_title()
        if not window_title:
            if excel_count == 1:
                self.display.update_status('1 Excel instance connected')
            else:
                self.display.update_status(f'{excel_count} Excel instances, none active')
            return

        if excel_count > 1:
            self.display.update_status(f'Tracking: {window_title} ({excel_count} Excel instances)')
        else:
            self.display.update_status(f'Tracking: {window_title}')

    def safe_quit(self) -> None:
        self._cancel_timers()
        self.scraper.release_resources()
        self.display.cleanup()

    def _cancel_timers(self) -> None:
        try:
            if self.check_timer:
                self.display.cancel_update(self.check_timer)
        except Exception as e:
            self.logger.error(f'Error canceling timers: {e}')

    def run(self) -> None:
        try:
            self.display.start()
        except Exception as e:
            self.logger.error(f'Error running application: {e}')
            self.safe_quit()
