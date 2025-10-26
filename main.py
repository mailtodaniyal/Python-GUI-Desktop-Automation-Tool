import os, sys, time, threading, traceback, queue
from dataclasses import dataclass
from datetime import datetime
from typing import List
import numpy as np
from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QTextEdit, QFileDialog, QLabel, QLineEdit, QSpinBox, QDoubleSpinBox, QCheckBox, QFormLayout, QGroupBox, QTabWidget, QHBoxLayout
import pyautogui
import sounddevice as sd
try:
    import vlc
except:
    vlc = None
try:
    from pyexcel_ods3 import get_data as ods_get_data, save_data as ods_save_data
except:
    ods_get_data = None
    ods_save_data = None
IS_WIN = os.name == "nt"

@dataclass
class RowEntry:
    phone: str
    audio_path: str
    status: str = ""
    last_updated: str = ""

@dataclass
class AppConfig:
    dial_click_x: int = 0
    dial_click_y: int = 0
    hang_click_x: int = 0
    hang_click_y: int = 0
    use_auto_hang: bool = True
    audio_backend: str = "VLC"
    dtmf_target: str = "2233"
    dtmf_timeout_s: float = 30.0
    wait_after_dial_s: float = 2.0
    wait_before_audio_s: float = 1.0
    call_timeout_s: float = 30.0
    loop_delay_s: float = 1.5
    spreadsheet_path: str = ""

class GuiLogger(QObject):
    log_signal = Signal(str)
    def __init__(self):
        super().__init__()
        self._buffer = []
    def log(self, msg: str):
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] {msg}"
        self._buffer.append(line)
        self.log_signal.emit(line)

class SpreadsheetHandler:
    def __init__(self, logger: GuiLogger):
        self.logger = logger
        self.path = ""
        self.rows: List[RowEntry] = []
    def load_ods(self, path: str):
        if ods_get_data is None:
            raise RuntimeError("pyexcel-ods3 not installed")
        data = ods_get_data(path)
        if not data: raise RuntimeError("Empty ODS")
        sheet = None
        for k,v in data.items():
            if v: sheet=v; break
        processed=[]
        for r in sheet:
            if not r: continue
            processed.append([str(c).strip() for c in r])
        if processed and set(map(str.lower, processed[0])) & {"phone","audio"}:
            processed=processed[1:]
        rows=[]
        for r in processed:
            phone=r[0] if len(r)>0 else ""
            audio=r[1] if len(r)>1 else ""
            status=r[2] if len(r)>2 else ""
            last=r[3] if len(r)>3 else ""
            if phone and audio: rows.append(RowEntry(phone,audio,status,last))
        if not rows: raise RuntimeError("No valid rows")
        self.path=path
        self.rows=rows
        self.logger.log(f"Loaded {len(rows)} rows from {path}")
    def save_ods(self):
        if ods_save_data is None: return
        data={"Sheet1":[["phone","audio","status","last_updated"]]}
        for r in self.rows:
            data["Sheet1"].append([r.phone,r.audio_path,r.status,r.last_updated])
        ods_save_data(self.path,data)

class AudioPlayer:
    def __init__(self, backend="VLC"):
        self.backend=backend
        self.vlc_player=None
    def play(self,path:str):
        if self.backend=="VLC" and vlc:
            self.vlc_player=vlc.MediaPlayer(path)
            self.vlc_player.play()
            time.sleep(0.5)
            while self.vlc_player.is_playing(): time.sleep(0.1)
        elif self.backend=="sounddevice":
            import soundfile as sf
            data,fs=sf.read(path)
            sd.play(data,fs)
            sd.wait()
        elif IS_WIN and self.backend=="winsound":
            import winsound
            winsound.PlaySound(path,winsound.SND_FILENAME)
        else:
            if sys.platform.startswith("darwin"): os.system(f"afplay '{path}'")
            elif sys.platform.startswith("linux"): os.system(f"xdg-open '{path}'")
            elif IS_WIN: os.startfile(path)

class CallWorker(QThread):
    log_signal=Signal(str)
    finished_signal=Signal()
    def __init__(self,config:AppConfig,rows:List[RowEntry]):
        super().__init__()
        self.config=config
        self.rows=rows
        self._stop=False
    def run(self):
        player=AudioPlayer(self.config.audio_backend)
        for row in self.rows:
            if self._stop: break
            try:
                self.log_signal.emit(f"Dialing {row.phone}")
                pyautogui.click(self.config.dial_click_x,self.config.dial_click_y)
                pyautogui.typewrite(row.phone)
                pyautogui.press("enter")
                time.sleep(self.config.wait_after_dial_s)
                time.sleep(self.config.wait_before_audio_s)
                self.log_signal.emit(f"Playing {row.audio_path}")
                player.play(row.audio_path)
                time.sleep(self.config.call_timeout_s)
                if self.config.use_auto_hang:
                    pyautogui.click(self.config.hang_click_x,self.config.hang_click_y)
                    self.log_signal.emit("Hung up")
                row.status="DONE"
                row.last_updated=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            except Exception as e:
                self.log_signal.emit("Error: "+str(e))
            time.sleep(self.config.loop_delay_s)
        self.finished_signal.emit()
    def stop(self):
        self._stop=True

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("desktopWksSpredsheet23")
        self.logger=GuiLogger()
        self.config=AppConfig()
        self.spread=SpreadsheetHandler(self.logger)
        self.worker=None
        self.tabs=QTabWidget()
        self.log_box=QTextEdit()
        self.log_box.setReadOnly(True)
        self.logger.log_signal.connect(self.log_box.append)
        self.btn_load=QPushButton("Load Spreadsheet")
        self.btn_start=QPushButton("Start")
        self.btn_stop=QPushButton("Stop")
        self.btn_load.clicked.connect(self.load_sheet)
        self.btn_start.clicked.connect(self.start_calls)
        self.btn_stop.clicked.connect(self.stop_calls)
        layout=QVBoxLayout()
        layout.addWidget(self.btn_load)
        layout.addWidget(self.btn_start)
        layout.addWidget(self.btn_stop)
        layout.addWidget(self.tabs)
        layout.addWidget(self.log_box)
        self.setLayout(layout)
        self.build_settings_tab()
    def build_settings_tab(self):
        w=QWidget()
        f=QFormLayout()
        self.x_dial=QSpinBox(); self.x_dial.setMaximum(9999)
        self.y_dial=QSpinBox(); self.y_dial.setMaximum(9999)
        self.x_hang=QSpinBox(); self.x_hang.setMaximum(9999)
        self.y_hang=QSpinBox(); self.y_hang.setMaximum(9999)
        f.addRow("Dial X",self.x_dial)
        f.addRow("Dial Y",self.y_dial)
        f.addRow("Hang X",self.x_hang)
        f.addRow("Hang Y",self.y_hang)
        w.setLayout(f)
        self.tabs.addTab(w,"Settings")
    def load_sheet(self):
        file,_=QFileDialog.getOpenFileName(self,"Open Spreadsheet","","ODS Files (*.ods)")
        if file:
            try:
                self.spread.load_ods(file)
                self.logger.log("Spreadsheet loaded")
            except Exception as e:
                QMessageBox.critical(self,"Error",str(e))
    def start_calls(self):
        self.config.dial_click_x=self.x_dial.value()
        self.config.dial_click_y=self.y_dial.value()
        self.config.hang_click_x=self.x_hang.value()
        self.config.hang_click_y=self.y_hang.value()
        if not self.spread.rows:
            QMessageBox.warning(self,"No data","Load spreadsheet first")
            return
        self.worker=CallWorker(self.config,self.spread.rows)
        self.worker.log_signal.connect(self.logger.log)
        self.worker.finished_signal.connect(self.calls_finished)
        self.worker.start()
    def stop_calls(self):
        if self.worker: self.worker.stop()
    def calls_finished(self):
        self.logger.log("Automation finished")
        try: self.spread.save_ods()
        except: pass

if __name__=="__main__":
    app=QApplication(sys.argv)
    w=MainWindow()
    w.resize(600,400)
    w.show()
    sys.exit(app.exec())
