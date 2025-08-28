# Next Button

from PyQt6.QtWidgets import QPushButton

def NextButton(label,Enable):
    next_button = QPushButton(label)
    next_button.setEnabled(Enable)
    if not Enable:
        next_button.setStyleSheet("background-color: #53575A; color: white;")
    else:
        next_button.setStyleSheet("background-color: #0052CC; color: white;")
    return next_button