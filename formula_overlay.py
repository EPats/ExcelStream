import tkinter as tk
import logging
from abc import ABC, abstractmethod

"""
Formula display handling, abstracted to allow
for a more adaptive response because I am going
to be trying different ways to create the overlay.

Essentially, do more work so in future there is less work.
#Relatable.
"""


# Constants
WINDOW_WIDTH = 400
WINDOW_HEIGHT = 200
WINDOW_OPACITY = 0.9
TEXT_WRAP_WIDTH = 380
FORMULA_FONT = ('Consolas', 10)
STATUS_FONT = ('Consolas', 8)


class FormulaDisplayBase(ABC):
    """
    Abstract base class for formula display implementation,
    so we can swap out different visualisation approaches.
    """

    @abstractmethod
    def show(self) -> None:
        """Show the display"""
        pass

    @abstractmethod
    def hide(self) -> None:
        """Hide the display"""
        pass

    @abstractmethod
    def update_formula(self, formula_data: dict[str, str] | None) -> None:
        """Update the formula displayed"""
        pass

    @abstractmethod
    def update_status(self, status_text: str) -> None:
        """Update the status text"""
        pass

    @abstractmethod
    def set_error(self, error_message: str) -> None:
        """Display an error message"""
        pass

    @abstractmethod
    def schedule_update(self, callback, ms: int):
        """Schedule a callback to be run after a delay"""
        pass

    @abstractmethod
    def cancel_update(self, timer_id) -> None:
        """Cancel a previously scheduled update"""
        pass

    @abstractmethod
    def cleanup(self) -> None:
        """Release resources and prepare for shutdown"""
        pass

    @abstractmethod
    def start(self) -> None:
        """Start the main loop of the display"""
        pass

    @abstractmethod
    def set_close_handler(self, handler) -> None:
        """Set a handler to be called when the display is closed"""
        pass


class TkinterFormulaDisplay(FormulaDisplayBase):
    """
    Tkinter implementation of the formula display.
    This is the first approach I used.
    It probably won't be the final one I use.

    Creates a floating overlay window showing formula information.
    """

    def __init__(self) -> None:
        self.logger = logging.getLogger(__name__)
        self._setup_ui()

    def _setup_ui(self) -> None:

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

    def show(self) -> None:
        """Make the window visible"""
        self.root.deiconify()

    def hide(self) -> None:
        """Hide the window"""
        self.root.withdraw()

    def update_formula(self, formula_data: dict[str, str] | None) -> None:
        """Update the formula displayed in the window"""
        if formula_data:
            display_text = f"{formula_data['address']}:\n{formula_data['formula']}"
            self.formula_label.config(text=display_text)
        else:
            self.formula_label.config(text='No formula detected')

    def update_status(self, status_text: str) -> None:
        """Update the status text at the bottom of the window"""
        self.instance_label.config(text=status_text)

    def set_error(self, error_message: str) -> None:
        """Display an error message in the formula area"""
        self.formula_label.config(text=f'Error: {error_message}')

    def schedule_update(self, callback, ms: int) -> str:
        """Schedule a callback to be run after a delay"""
        return self.root.after(ms, callback)

    def cancel_update(self, timer_id: str) -> None:
        """Cancel a previously scheduled update"""
        self.root.after_cancel(timer_id)

    def cleanup(self) -> None:
        """Cleanup resources before shutdown"""
        try:
            self.root.quit()
            self.root.destroy()
        except Exception as e:
            self.logger.error(f'Error closing window: {e}')

    def start(self) -> None:
        """Start the main tkinter event loop"""
        try:
            self.root.mainloop()
        except Exception as e:
            self.logger.error(f'Error in mainloop: {e}')
            self.cleanup()

    def set_close_handler(self, handler) -> None:
        """Set handler for window close event"""
        self.root.protocol('WM_DELETE_WINDOW', handler)


# Factory function to create the appropriate display
def create_formula_display(display_type: str = 'tkinter') -> FormulaDisplayBase:
    """
    Factory function to create the appropriate display type.

    Args:
        display_type: The type of display to create ('tkinter' or future options)

    Returns:
        An instance of FormulaDisplayBase
    """
    if display_type.lower() == 'tkinter':
        return TkinterFormulaDisplay()
    else:
        # Default to tkinter if unknown type
        return TkinterFormulaDisplay()