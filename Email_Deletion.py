import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton
import subprocess

def run_email_script():
    subprocess.call(['python', 'email_cleaner.py'])

# Create the main application window
app = QApplication(sys.argv)
window = QMainWindow()
window.setWindowTitle("Email Script")

# Create a button to run the email script
run_button = QPushButton("Run Email Script", window)
run_button.clicked.connect(run_email_script)
run_button.setGeometry(50, 50, 200, 50)

# Show the window
window.show()

# Start the application event loop
sys.exit(app.exec_())
