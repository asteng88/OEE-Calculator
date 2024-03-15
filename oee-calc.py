from PyQt6 import (uic, QtWidgets, QtGui)
from PyQt6.QtWidgets import (QMainWindow, QApplication, QMessageBox)
from PyQt6.QtGui import QIcon
from PyQt6.uic import loadUi
import os, sys

basedir = os.path.dirname(__file__)  # Set a path as the same directory as this file

# this block ensures that the ICON is definately displayed on the taskbar in Windows
try:
    from ctypes import windll

    myappid = 'com.asteng88.OEE-Calc.OEE-Calc.1.0'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass


class UiMainWindow(QMainWindow):
    def __init__(self):
        super(UiMainWindow, self).__init__()
        ui = os.path.join(basedir, 'ops_toolbox.ui')  # Set the path of the ui file
        uic.loadUi(ui, self)

        # Exit button action is handled in the ui file. No need to connect it here
        # Find widgets
        self.planned_time_input = self.findChild(QtWidgets.QLineEdit, 'planned_time_input')
        self.run_time_input = self.findChild(QtWidgets.QLineEdit, 'run_time_input')
        self.target_count_input = self.findChild(QtWidgets.QLineEdit, 'target_count_input')
        self.total_parts_made_input = self.findChild(QtWidgets.QLineEdit, 'total_parts_made_input')
        self.good_parts_input = self.findChild(QtWidgets.QLineEdit, 'good_parts_input')
        self.answers_box = self.findChild(QtWidgets.QTextBrowser, 'answers_box')

        # Connect calculate button to the slot
        self.calculate_button = self.findChild(QtWidgets.QPushButton, 'calculate_button')
        self.calculate_button.clicked.connect(self.calculate_oee)

        # Connect clear button to the slot
        self.clear_button = self.findChild(QtWidgets.QPushButton, 'clear_button')
        self.clear_button.clicked.connect(self.clear_contents)

    def calculate_oee(self):
        try:
            run_time = float(self.run_time_input.text())
            planned_production_time = float(self.planned_time_input.text())
            total_count = float(self.total_parts_made_input.text())
            target_count = float(self.target_count_input.text())
            good_count = float(self.good_parts_input.text())

            availability = run_time / planned_production_time
            performance = total_count / target_count
            quality = good_count / total_count

            oee = availability * performance * quality

            factors = {'Availability': availability, 'Performance': performance, 'Quality': quality}
            lowest_factor = min(factors, key=factors.get)

            result_text = (f'The OEE percentage is: {(oee * 100):.1f}%\n\n'
                           f'Availability: {availability * 100:.1f}%\n'
                           f'Performance: {performance * 100:.1f}%\n'
                           f'Quality: {quality * 100:.1f}%\n\n'
                           f'The area affecting OEE the most is: {lowest_factor}')

            self.answers_box.append(result_text)

        except ValueError:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Icon.Critical)
            msg.setText("Error")
            msg.setInformativeText('Please ensure all fields are complete and are numeric')
            msg.setWindowTitle("Error")
            msg.setWindowIcon(QIcon(icon_1))
            msg.exec()

    def clear_contents(self):
        # Clear all line edits and the text box
        self.planned_time_input.clear()
        self.run_time_input.clear()
        self.target_count_input.clear()
        self.total_parts_made_input.clear()
        self.good_parts_input.clear()
        self.answers_box.clear()

    def quit_application(self):
        # Exit the application
        self.close()


app = QApplication(sys.argv)
window = UiMainWindow()
icon_1 = os.path.join (basedir, 'calculator2.png')
window.setWindowIcon(QIcon(icon_1))
window.show()
app.exec()
