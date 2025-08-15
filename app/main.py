import sys
from PyQt6.QtWidgets import QApplication
from gui.main_gui import MainApp


def main():
    """
    Main entry point of the application.
    Creates a QApplication instance, launches the MainApp (main window),
    and starts the PyQt6 event loop.
    """
    # Initialize the Qt application
    app = QApplication(sys.argv)

    # Create and show the main application window
    window = MainApp()
    window.show()

    # Execute the Qt event loop until the application is closed
    sys.exit(app.exec())


if __name__ == "__main__":
    # Run only if this file is executed directly (not imported as a module)
    main()