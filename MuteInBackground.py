import sys
import psutil
import win32gui
import win32process
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QListWidget, QListWidgetItem, QLabel, QSystemTrayIcon, QMenu, QAction, QCheckBox
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QIcon
from threading import Thread
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

    def unmute_all(self):
        for app_name in list(self.muted_apps.keys()):
            self.unmute_app(app_name)


class App(QWidget):
    def __init__(self):
        super().__init__()

        self.tray_icon_enabled = True
        self.show_all_apps = False
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

        self.tray_checkbox = QCheckBox("Enable tray icon")
        self.tray_checkbox.setChecked(True)
        self.tray_checkbox.stateChanged.connect(self.toggle_tray_icon)
        right_layout.addWidget(self.tray_checkbox)

        self.show_all_checkbox = QCheckBox("Show all applications")
        self.show_all_checkbox.setChecked(False)
        self.show_all_checkbox.stateChanged.connect(self.toggle_show_all_apps)
        right_layout.addWidget(self.show_all_checkbox)

        main_layout.addLayout(left_layout)
        main_layout.addLayout(center_layout)
        main_layout.addLayout(right_layout)

        self.setLayout(main_layout)

        self.setWindowTitle('Mute In Background')
        self.setGeometry(300, 300, 600, 400)

        # Set the window icon
        self.setWindowIcon(QIcon("MuteInBackground.png"))

        self.timer = QTimer()
        self.timer.timeout.connect(self.check_focus)
        self.timer.start(400)  # Check focus every 400ms

        self.create_tray_icon()

    def create_tray_icon(self):
        try:
            print("Creating tray icon...")
            icon = QIcon("MuteInBackground.png")  # Ensure you have MuteInBackground.ico in the same directory
            if not icon.isNull():
                print("Icon loaded successfully")
            else:
                print("Failed to load icon")

            self.tray_icon = QSystemTrayIcon(icon, self)
            self.tray_icon.setToolTip("Mute In Background")

            self.tray_icon.activated.connect(self.on_tray_icon_activated)

            tray_menu = QMenu()
            show_action = QAction("Show", self)
            quit_action = QAction("Quit", self)

            show_action.triggered.connect(self.show_window)
            quit_action.triggered.connect(QApplication.instance().quit)

            tray_menu.addAction(show_action)
            tray_menu.addAction(quit_action)

            self.tray_icon.setContextMenu(tray_menu)
            self.tray_icon.show()

            if self.tray_icon.isVisible():
                print("Tray icon is visible")
            else:
                print("Tray icon is not visible")
        except Exception as e:
            print(f"Error creating tray icon: {e}")

    def on_tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.show_window()

    def hide_window(self):
        self.hide()
        if self.tray_icon_enabled:
            self.tray_thread = Thread(target=self.create_tray_icon)
            self.tray_thread.daemon = True
            self.tray_thread.start()

    def show_window(self):
        self.show()
        self.activateWindow()

    def closeEvent(self, event):
        print("Unmuting all applications before closing...")
        self.audio_manager.unmute_all()

        if self.tray_icon_enabled:
            event.ignore()
            self.hide_window()
            self.tray_icon.showMessage(
                "Mute In Background",
                "Application minimized to tray.",
                QSystemTrayIcon.Information,
                2000
            )
            print("Application minimized to tray")
        else:
            event.accept()
            print("Application closed")

    def refresh_app_list(self):
        self.untracked_list.clear()
        apps = {}
        tracked_apps = {self.tracked_list.item(index).text() for index in range(self.tracked_list.count())}

        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] not in tracked_apps and not proc.name().startswith('svchost') and not proc.username().endswith('SYSTEM'):
                    if self.show_all_apps or self.is_user_facing_app(proc.info['pid']):
                        app_name = self.get_window_title(proc.info['pid'], proc.info['name'])
                        if app_name and app_name not in apps:
                            apps[app_name] = proc.info['name']
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue

        for app_name, exe_name in sorted(apps.items()):
            item = QListWidgetItem(f"{app_name} ({exe_name})")
            item.setData(Qt.UserRole, exe_name)
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
            self.audio_manager.unmute_app(item.data(Qt.UserRole))
        self.refresh_app_list()

    def check_focus(self):
        active_window = self.get_active_window_name()
        for index in range(self.tracked_list.count()):
            item = self.tracked_list.item(index)
            app_name = item.data(Qt.UserRole)
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

    def get_window_title(self, pid, exe_name):
        def callback(hwnd, pid):
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            if found_pid == pid:
                title = win32gui.GetWindowText(hwnd)
                if win32gui.IsWindowVisible(hwnd) and title:
                    return title
            return None

        window_titles = []
        win32gui.EnumWindows(lambda hwnd, resultList: resultList.append(callback(hwnd, pid)), window_titles)
        window_titles = [title for title in window_titles if title and not title.startswith("Microsoft Text Input Application")]

        if window_titles:
            title = window_titles[0]

            if "Firefox" in title:
                return "Firefox"
            # Fallback to executable name if title contains dynamic content like song names or URLs
            if " - " in title or "http://" in title or "https://" in title:
                return exe_name
            return title
        return exe_name

    def is_user_facing_app(self, pid):
        try:
            for hwnd in self.enum_windows_for_pid(pid):
                if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                    return True
        except Exception as e:
            print(f"Error checking user-facing app: {e}")
        return False

    def enum_windows_for_pid(self, pid):
        def callback(hwnd, hwnds):
            if win32process.GetWindowThreadProcessId(hwnd)[1] == pid:
                hwnds.append(hwnd)
            return True

        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        return hwnds

    def toggle_tray_icon(self, state):
        self.tray_icon_enabled = (state == Qt.Checked)
        if not self.tray_icon_enabled:
            if hasattr(self, 'tray_icon') and self.tray_icon.isVisible():
                self.tray_icon.hide()
        else:
            if hasattr(self, 'tray_icon') and not self.tray_icon.isVisible():
                self.tray_icon.show()

    def toggle_show_all_apps(self, state):
        self.show_all_apps = (state == Qt.Checked)
        self.refresh_app_list()


def main():
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
