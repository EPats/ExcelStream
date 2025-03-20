import tkinter as tk
import logging
from abc import ABC, abstractmethod
import customtkinter as ctk

"""
Formula display handling, abstracted to allow
for a more adaptive response because I am going
to be trying different ways to create the overlay.

Essentially, do more work so in future there is less work.
#Relatable.
"""


def get_screen_dimensions():
    """Get the screen dimensions using tkinter"""
    temp_root = tk.Tk()
    screen_width = temp_root.winfo_screenwidth()
    screen_height = temp_root.winfo_screenheight()
    temp_root.destroy()
    return screen_width, screen_height


# Get screen dimensions
SCREEN_WIDTH, SCREEN_HEIGHT = get_screen_dimensions()

# Constants
WINDOW_WIDTH = int(SCREEN_WIDTH * 0.95)
WINDOW_HEIGHT = 200
WINDOW_OPACITY = 0.9
TEXT_WRAP_WIDTH = int(WINDOW_WIDTH * 0.95)
FORMULA_FONT = ('Roboto', 30)
STATUS_FONT = ('Roboto', 25)


class FormulaDisplayBase(ABC):
    """
    Abstract base class for formula display implementation,
    so we can swap out different visualisation approaches.
    """

    @abstractmethod
    def __init__(self, debug_mode: bool = False) -> None:
        """Initialize the display with optional debug mode"""
        pass

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

    def __init__(self, debug_mode: bool = False) -> None:
        self.logger = logging.getLogger(__name__)
        self.debug_mode: bool = debug_mode
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

        # Status display (only shown in debug mode)
        self.instance_label: tk.Label = tk.Label(
            self.root,
            text='Searching for Excel instances...',
            font=STATUS_FONT
        )

        if self.debug_mode:
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
        """Update the status text at the bottom of the window (if in debug mode) """
        if self.debug_mode and hasattr(self, 'instance_label'):
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


class CustomTkinterFormulaDisplay(FormulaDisplayBase):
    """
    CustomTkinter implementation of the formula display.
    Creates a modern, floating overlay window showing formula information.
    """

    def __init__(self, debug_mode: bool = False) -> None:
        self.logger = logging.getLogger(__name__)
        self.debug_mode: bool = debug_mode

        # Set the appearance mode and default color theme
        ctk.set_appearance_mode("system")  # Options: "system", "dark", "light"
        ctk.set_default_color_theme("blue")  # Options: "blue", "green", "dark-blue"

        self._setup_ui()

    def _make_window_draggable(self) -> None:
        """Make the window draggable by clicking and dragging anywhere"""

        def start_drag(event):
            self._drag_x = event.x
            self._drag_y = event.y

        def do_drag(event):
            x = self.root.winfo_x() + event.x - self._drag_x  # type: ignore
            y = self.root.winfo_y() + event.y - self._drag_y  # type: ignore
            self.root.geometry(f"+{x}+{y}")

        # Bind the events to the main window and all its children
        self.root.bind("<ButtonPress-1>", start_drag)
        self.root.bind("<B1-Motion>", do_drag)

    def _setup_ui(self) -> None:
        # Main window
        self.root = ctk.CTk()
        self.root.title("Excel Formula Live Tracker")
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", WINDOW_OPACITY)
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")

        # Remove window decoration (minimize/maximize/close buttons)
        self.root.overrideredirect(True)

        # Add draggable functionality since we removed the title bar
        self._make_window_draggable()

        # Create a frame for better organization
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill=ctk.BOTH, expand=True, padx=10, pady=10)

        # Header frame for close button only
        self.header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.header_frame.pack(fill=ctk.X, padx=10, pady=(5, 0))

        # Close button
        self.close_button = ctk.CTkButton(
            self.header_frame,
            text="×",
            width=20,
            height=20,
            corner_radius=10,
            fg_color="transparent",
            hover_color=("gray75", "gray25"),
            command=self.root.destroy
        )
        self.close_button.pack(side=ctk.RIGHT, anchor="e")

        # Formula address label
        self.address_label = ctk.CTkLabel(
            self.main_frame,
            text="Waiting for Excel...",
            font=FORMULA_FONT
        )
        self.address_label.pack(padx=10, pady=(5, 0), anchor="w", fill="x")

        # Create a container frame for formula and status (if in debug mode)
        self.content_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.content_frame.pack(fill=ctk.BOTH, expand=True, padx=5, pady=5)

        # Status display (always create it, but only show in debug mode)
        self.status_frame = ctk.CTkFrame(self.content_frame, fg_color=("gray90", "gray20"), corner_radius=5)
        self.instance_label = ctk.CTkLabel(
            self.status_frame,
            text="Searching for Excel instances...",
            font=STATUS_FONT,
            text_color=("gray50", "gray70")
        )
        self.instance_label.pack(pady=2, padx=5, fill=ctk.X)

        # Only show the status frame if debug mode is enabled
        if self.debug_mode:
            self.status_frame.pack(fill=ctk.X, pady=(0, 5), padx=5)

        # Formula display - use a text widget for better formatting
        self.formula_text = ctk.CTkTextbox(
            self.content_frame,
            height=100,  # Adjusted height to make room for debug text
            font=FORMULA_FONT,
            wrap="word"
        )
        self.formula_text.pack(padx=5, pady=5, fill=ctk.BOTH, expand=True)
        self.formula_text.insert("1.0", "No formula detected")
        self.formula_text.configure(state="disabled")  # Make it read-only

        # Status display is now handled in the content_frame

    def show(self) -> None:
        """Make the window visible"""
        self.root.deiconify()

    def hide(self) -> None:
        """Hide the window"""
        self.root.withdraw()

    def update_formula(self, formula_data: dict[str, str] | None) -> None:
        """Update the formula displayed in the window"""
        if formula_data:
            # Update address label
            self.address_label.configure(text=formula_data['address'])

            # Update formula text
            self.formula_text.configure(state="normal")
            self.formula_text.delete("1.0", "end")
            self.formula_text.insert("1.0", formula_data['formula'])
            self.formula_text.configure(state="disabled")
        else:
            self.address_label.configure(text="No cell selected")
            self.formula_text.configure(state="normal")
            self.formula_text.delete("1.0", "end")
            self.formula_text.insert("1.0", "No formula detected")
            self.formula_text.configure(state="disabled")

    def update_status(self, status_text: str) -> None:
        """Update the status text at the bottom of the window (if in debug mode) """
        # Always update the text, but only show if in debug mode
        if hasattr(self, 'instance_label'):
            self.instance_label.configure(text=status_text)

            # Make sure the status frame is properly shown/hidden based on debug mode
            if hasattr(self, 'status_frame'):
                if self.debug_mode:
                    # Check if it's already packed
                    if not self.status_frame.winfo_ismapped():
                        self.status_frame.pack(fill=ctk.X, pady=(0, 5), padx=5, before=self.formula_text)
                else:
                    # Hide it if not in debug mode
                    if self.status_frame.winfo_ismapped():
                        self.status_frame.pack_forget()

    def set_error(self, error_message: str) -> None:
        """Display an error message in the formula area"""
        self.address_label.configure(text="Error")
        self.formula_text.configure(state="normal")
        self.formula_text.delete("1.0", "end")
        self.formula_text.insert("1.0", f"Error: {error_message}")
        self.formula_text.configure(state="disabled")

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
            self.logger.error(f"Error closing window: {e}")

    def start(self) -> None:
        """Start the main tkinter event loop"""
        try:
            self.root.mainloop()
        except Exception as e:
            self.logger.error(f"Error in mainloop: {e}")
            self.cleanup()

    def set_close_handler(self, handler) -> None:
        """Set handler for window close event"""
        self.root.protocol("WM_DELETE_WINDOW", handler)


def create_formula_display(display_type: str = 'tkinter', debug_mode: bool = False) -> FormulaDisplayBase:
    """ Factory function to create the appropriate display type. """

    if display_type.lower() == 'tkinter':
        return TkinterFormulaDisplay(debug_mode=debug_mode)
    elif display_type.lower() == "customtkinter":
        return CustomTkinterFormulaDisplay(debug_mode=debug_mode)
    else:
        # Default to customtkinter if unknown type
        return CustomTkinterFormulaDisplay(debug_mode=debug_mode)
