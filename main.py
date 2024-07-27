import sys
import psutil
import win32gui
import win32process
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QListWidget, QListWidgetItem, QCheckBox
from PyQt5.QtCore import QTimer, Qt
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume

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
        self.muted_apps = []

    def initUI(self):
        layout = QVBoxLayout()

        self.listWidget = QListWidget()
        self.refresh_button = QPushButton("Refresh Applications")
        self.refresh_button.clicked.connect(self.refresh_app_list)

        self.check_button = QPushButton("Check Focus")
        self.check_button.clicked.connect(self.check_focus)

        layout.addWidget(self.refresh_button)
        layout.addWidget(self.listWidget)
        layout.addWidget(self.check_button)

        self.setLayout(layout)

        self.setWindowTitle('Application Audio Manager')
        self.setGeometry(300, 300, 300, 400)

        self.timer = QTimer()
        self.timer.timeout.connect(self.check_focus)
        self.timer.start(200)  # Check focus every second

    def refresh_app_list(self):
        self.listWidget.clear()
        apps = []
        for proc in psutil.process_iter(['pid', 'name']):
            apps.append((proc.info['name'], proc.info['pid']))

        # Sort apps alphabetically by name
        apps.sort(key=lambda x: x[0])

        for app_name, pid in apps:
            item = QListWidgetItem(app_name)
            item.setData(Qt.UserRole, pid)
            item.setCheckState(Qt.Unchecked)
            self.listWidget.addItem(item)

    def check_focus(self):
        active_window = self.get_active_window_name()
        for index in range(self.listWidget.count()):
            item = self.listWidget.item(index)
            app_name = item.text()
            if item.checkState() == Qt.Checked:
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
