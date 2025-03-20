"""
Excel Formula Tracker

Creates a floating window that tracks the most recent formula used/highlighted
"""

import logging
import argparse
from excel_formula_tracker import ExcelFormulaTracker


def setup_logging(verbose: bool = False) -> None:
    """Configure logging for the application"""
    log_level = logging.DEBUG if verbose else logging.INFO

    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )


def parse_arguments():
    parser = argparse.ArgumentParser(description='Excel Formula Tracker')

    parser.add_argument(
        '--display',
        choices=['tkinter', 'customtkinter'],
        default='customtkinter',
        help='Display type to use (default: customtkinter)'
    )

    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose logging'
    )

    parser.add_argument(
        '--debug',
        action='store_true',
        help='Enable debug mode with additional status information'
    )

    return parser.parse_args()


def main():
    args = parse_arguments()
    setup_logging(args.verbose)

    tracker = ExcelFormulaTracker(
        display_type=args.display,
        debug_mode=args.debug
    )
    tracker.run()


if __name__ == '__main__':
    main()
