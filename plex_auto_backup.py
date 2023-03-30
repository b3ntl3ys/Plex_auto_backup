import os
import sys
import psutil
import zipfile
import winreg
import subprocess
from datetime import datetime, timedelta, time
from PyQt5.QtCore import QSettings, QTimer, QThread, pyqtSignal, Qt,QTime
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QPushButton, QLabel, QVBoxLayout, QSpinBox, QTextEdit, QCheckBox, QMessageBox, QTimeEdit


class FolderZipper(QWidget):
    def __init__(self):
        super().__init__()
        self.settings = QSettings('FolderZipper', 'FolderZipper')
        self.folder_path = self.settings.value('folder_path', '')
        self.save_path = self.settings.value('save_path', '')
        self.interval = int(self.settings.value('interval', '7'))
        self.export_plex_registry_key = self.settings.value('export_plex_registry_key', False, type=bool)

        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('Folder Zipper')
        self.setFixedSize(400, 400)
        layout = QVBoxLayout()

        start_with_windows_checkbox = QCheckBox('Start with Windows')
        start_with_windows_checkbox.setChecked(self.settings.value('start_with_windows', False, type=bool))
        start_with_windows_checkbox.stateChanged.connect(self.update_start_with_windows)
        layout.addWidget(start_with_windows_checkbox)

        export_plex_registry_key_checkbox = QCheckBox('Export Plex Media Server registry key')
        export_plex_registry_key_checkbox.setChecked(self.export_plex_registry_key)
        export_plex_registry_key_checkbox.stateChanged.connect(self.update_export_plex_registry_key_setting)
        layout.addWidget(export_plex_registry_key_checkbox)

        folder_label = QLabel('Select a folder to zip:')
        layout.addWidget(folder_label)

        folder_button = QPushButton('Choose Folder')
        folder_button.clicked.connect(self.choose_folder)
        layout.addWidget(folder_button)

        folder_path_edit = QTextEdit(self.folder_path)
        folder_path_edit.setReadOnly(True)
        layout.addWidget(folder_path_edit)

        save_label = QLabel('Select where to save the zip file:')
        layout.addWidget(save_label)

        save_button = QPushButton('Choose Save Location')
        save_button.clicked.connect(self.choose_save_location)
        layout.addWidget(save_button)

        save_path_edit = QTextEdit(self.save_path)
        save_path_edit.setReadOnly(True)
        layout.addWidget(save_path_edit)

        zip_button = QPushButton('Zip Folder')
        zip_button.clicked.connect(self.start_zip_thread)
        layout.addWidget(zip_button)

        backup_time_label = QLabel('Select the time of day to backup the folder:')
        layout.addWidget(backup_time_label)

        backup_time_edit = QTimeEdit()
        backup_time_edit.setDisplayFormat('hh:mm')
        backup_time_edit.setTime(self.settings.value('backup_time', QTime(0, 0)))
        backup_time_edit.timeChanged.connect(self.update_backup_time)
        layout.addWidget(backup_time_edit)

        interval_label = QLabel('Zip folder every X days:')
        layout.addWidget(interval_label)

        interval_spinbox = QSpinBox()
        interval_spinbox.setMinimum(1)
        interval_spinbox.setMaximum(365)
        interval_spinbox.setValue(self.interval)
        layout.addWidget(interval_spinbox)
        self.interval_spinbox = interval_spinbox

        countdown_label = QLabel('Next zipping in {} days, {} hours, {} minutes, {} seconds'.format(self.interval, 0, 0, 0))
        self.countdown_label = countdown_label
        interval_spinbox.valueChanged.connect(self.update_interval)
        
        self.next_zip_time = datetime.now() + timedelta(days=self.interval)
        self.countdown_timer = QTimer()
        self.countdown_timer.timeout.connect(self.update_countdown)
        self.countdown_timer.start(1000)  
        layout.addWidget(countdown_label)
        self.countdown_label = countdown_label

        self.setLayout(layout)
        self.show()
        self.folder_path_edit = folder_path_edit
        self.save_path_edit = save_path_edit
        self.start_with_windows_checkbox = start_with_windows_checkbox
        self.backup_time_edit = backup_time_edit

        self.countdown_timer.stop()
        self.countdown_timer.start(1000)
        

    def update_backup_time(self, time):
        self.settings.setValue('backup_time', time)
        backup_time = time.toPyTime()
        self.next_zip_time = self.get_next_backup_time()
        self.countdown_timer.stop()
        self.countdown_timer.start(1000)

    def get_next_backup_time(self):
        backup_time = self.backup_time_edit.time().toPyTime() 
        next_backup_time = datetime.combine(datetime.today(), backup_time)
        if next_backup_time <= datetime.now():
       
            next_backup_time += timedelta(days=1)
        return next_backup_time
    
    def update_interval(self, value):
        self.interval = value
        self.settings.setValue('interval', self.interval)
        self.next_zip_time = datetime.now() + timedelta(days=self.interval)
        self.countdown_timer.start(1000)  
        self.update_countdown()  
    def update_start_with_windows(self, state):
        self.settings.setValue('start_with_windows', state == Qt.Checked)

    def update_export_plex_registry_key_setting(self):
        self.export_plex_registry_key = self.sender().isChecked()
        self.settings.setValue('export_plex_registry_key', self.export_plex_registry_key)

    def choose_folder(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ShowDirsOnly
        folder_path = QFileDialog.getExistingDirectory(self, "Select a folder to zip", options=options)
        if folder_path:
            self.folder_path = folder_path
            self.settings.setValue('folder_path', self.folder_path)
            self.folder_path_edit.setText(self.folder_path)

    def choose_save_location(self):
        save_path, _ = QFileDialog.getSaveFileName(self, "Select where to save the zip file", self.save_path, "Zip files (*.zip)")
        if save_path:
            self.save_path = save_path
            self.settings.setValue('save_path', self.save_path)
            self.save_path_edit.setText(self.save_path)

    def start_zip_thread(self):
        self.kill_task_by_name()

        current_datetime_str = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        save_path_parts = os.path.splitext(self.save_path)
        save_path_with_datetime = f"{save_path_parts[0]}_{current_datetime_str}{save_path_parts[1]}"

        self.zip_thread = ZipThread(self.folder_path, save_path_with_datetime)
        self.zip_thread.finished.connect(self.zip_thread_finished)  
        self.zip_thread.start()

        if self.export_plex_registry_key:
            registry_key = r"HKEY_CURRENT_USER\Software\Plex, Inc.\Plex Media Server"
            try:
                subprocess.check_output(['reg', 'export', registry_key, os.path.join(self.folder_path, 'PlexMediaServer.reg')])
            except subprocess.CalledProcessError as e:
                message = f"Error exporting Plex Media Server registry key: {e}. The registry key may not exist."
                QMessageBox.warning(self, 'Error Exporting Registry Key', message, QMessageBox.Ok)

    def update_countdown(self):
        time_left = self.next_zip_time - datetime.now()
        total_seconds_left = time_left.total_seconds()
        days, remainder = divmod(total_seconds_left, 86400)
        hours, remainder = divmod(remainder, 3600)
        minutes, seconds = divmod(remainder, 60)
        seconds_str = '{:02d}'.format(int(seconds))
        self.countdown_label.setText('Next zipping in {} days, {} hours, {} minutes, {} seconds'.format(days, hours, minutes, seconds_str))

        if time_left.total_seconds() <= 0:
            self.interval = self.interval_spinbox.value()
            self.start_zip_thread()
            self.next_zip_time = datetime.now() + timedelta(days=self.interval)
            self.countdown_timer.start(1000)

    def kill_task_by_name(task_name):
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] == task_name:
                proc.kill()

    # terminate a running task using subprocess
    kill_task_by_name("Plex Media Server.exe")

    # start a task uses subprocess when zipping is done
    def zip_thread_finished(self):
        subprocess.Popen(['Plex Media Server.exe'])

class ZipThread(QThread):
    finished = pyqtSignal()

    def __init__(self, folder_path, save_path):
        super().__init__()
        self.folder_path = folder_path
        self.save_path = save_path

    def run(self):
        with zipfile.ZipFile(self.save_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(self.folder_path):
                for file in files:
                    zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), os.path.join(self.folder_path, '..')))
        self.finished.emit()

if __name__ == '__main__':
    app = QApplication([])
    folder_zipper = FolderZipper()

    # To Simulate countdown remove # below

    #folder_zipper.interval = 1
    #folder_zipper.next_zip_time = datetime.now() + timedelta(seconds=5)
    #folder_zipper.update_countdown()
    
    app.exec_()

    if folder_zipper.settings.value('start_with_windows', False, type=bool):
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Run', 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, 'FolderZipper', 0, winreg.REG_SZ, sys.executable)
        winreg.CloseKey(key)
