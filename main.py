import sys
import psutil
import win32gui
import win32process
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QListWidget, QListWidgetItem, \
    QLabel
from PyQt5.QtCore import QTimer, Qt
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume


class AudioManager:
    def __init__(self):
        self.muted_apps = {}

    def mute_app(self, app_name):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            if session.Process and session.Process.name() == app_name:
                volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                volume.SetMute(1, None)
                self.muted_apps[app_name] = session
                print(f"Muted {app_name}")

    def unmute_app(self, app_name):
        if app_name in self.muted_apps:
            session = self.muted_apps[app_name]
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            volume.SetMute(0, None)
            print(f"Unmuted {app_name}")
            del self.muted_apps[app_name]


class App(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()
        self.audio_manager = AudioManager()

    def initUI(self):
        main_layout = QHBoxLayout()

        left_layout = QVBoxLayout()
        self.untracked_label = QLabel("Running Applications")
        self.untracked_list = QListWidget()
        self.refresh_button = QPushButton("Refresh Applications")
        self.refresh_button.clicked.connect(self.refresh_app_list)
        left_layout.addWidget(self.untracked_label)
        left_layout.addWidget(self.refresh_button)
        left_layout.addWidget(self.untracked_list)

        center_layout = QVBoxLayout()
        self.track_button = QPushButton(">")
        self.track_button.clicked.connect(self.track_application)
        self.untrack_button = QPushButton("<")
        self.untrack_button.clicked.connect(self.untrack_application)
        center_layout.addStretch()
        center_layout.addWidget(self.track_button)
        center_layout.addWidget(self.untrack_button)
        center_layout.addStretch()

        right_layout = QVBoxLayout()
        self.tracked_label = QLabel("Applications that will be muted while not focused")
        self.tracked_list = QListWidget()
        right_layout.addWidget(self.tracked_label)
        right_layout.addWidget(self.tracked_list)

        main_layout.addLayout(left_layout)
        main_layout.addLayout(center_layout)
        main_layout.addLayout(right_layout)

        self.setLayout(main_layout)

        self.setWindowTitle('Application Audio Manager')
        self.setGeometry(300, 300, 600, 400)

        self.timer = QTimer()
        self.timer.timeout.connect(self.check_focus)
        self.timer.start(1000)  # Check focus every second

    def refresh_app_list(self):
        self.untracked_list.clear()
        apps = []
        tracked_apps = {self.tracked_list.item(index).text() for index in range(self.tracked_list.count())}

        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] not in tracked_apps and not proc.name().startswith(
                        'svchost') and not proc.username().endswith('SYSTEM'):
                    apps.append((proc.info['name'], proc.info['pid']))
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue

        # Sort apps alphabetically by name
        apps.sort(key=lambda x: x[0])

        for app_name, pid in apps:
            item = QListWidgetItem(app_name)
            item.setData(Qt.UserRole, pid)
            self.untracked_list.addItem(item)

    def track_application(self):
        selected_items = self.untracked_list.selectedItems()
        for item in selected_items:
            self.untracked_list.takeItem(self.untracked_list.row(item))
            self.tracked_list.addItem(item)

    def untrack_application(self):
        selected_items = self.tracked_list.selectedItems()
        for item in selected_items:
            self.tracked_list.takeItem(self.tracked_list.row(item))
            self.untracked_list.addItem(item)
            self.audio_manager.unmute_app(item.text())
        self.refresh_app_list()

    def check_focus(self):
        active_window = self.get_active_window_name()
        for index in range(self.tracked_list.count()):
            item = self.tracked_list.item(index)
            app_name = item.text()
            if app_name != active_window:
                self.audio_manager.mute_app(app_name)
            else:
                self.audio_manager.unmute_app(app_name)

    def get_active_window_name(self):
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['pid'] == pid:
                return proc.info['name']
        return None


def main():
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
