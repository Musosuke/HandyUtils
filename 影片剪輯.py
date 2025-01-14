import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QLabel, QVBoxLayout, QHBoxLayout, QPushButton, QWidget, QSlider, QLineEdit
)
from PyQt5.QtGui import QFont, QImage, QPixmap
from PyQt5.QtCore import Qt, QTimer
import cv2
import subprocess


def trim_video(input_video, start_frame, end_frame, output_video):
    try:
        command = [
            "ffmpeg", "-i", input_video,
            "-vf", f"select=between(n\,{start_frame}\,{end_frame})",
            "-vsync", "vfr",
            output_video
        ]
        subprocess.run(command, check=True)
        return output_video  # 成功時返回輸出的檔案路徑
    except subprocess.CalledProcessError as e:
        print(f"影片剪輯失敗：{e}")
        return None


def open_in_explorer(file_path):
    try:
        abs_path = os.path.abspath(file_path)
        subprocess.run(f'explorer /select,"{abs_path}"', shell=True)
    except Exception as e:
        print(f"無法開啟檔案總管：{e}")


class VideoTrimWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("影片剪輯工具")
        self.setGeometry(100, 100, 800, 600)

        self.video_path = None
        self.total_frames = 0
        self.current_frame = 0
        self.cap = None
        self.is_playing = False  # 播放狀態變數

        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # Video display label
        self.video_label = QLabel("請拖入影片檔案", self)
        self.video_label.setAlignment(Qt.AlignCenter)
        self.video_label.setStyleSheet("border: 1px solid black; background-color: #000;")
        self.video_label.setFixedHeight(400)
        layout.addWidget(self.video_label)

        # Slider and input fields
        control_layout = QHBoxLayout()

        self.slider = QSlider(Qt.Horizontal, self)
        self.slider.setEnabled(False)
        self.slider.sliderMoved.connect(self.update_frame_from_slider)
        control_layout.addWidget(self.slider)

        self.frame_input = QLineEdit(self)
        self.frame_input.setFixedWidth(80)
        self.frame_input.setAlignment(Qt.AlignCenter)
        self.frame_input.setEnabled(False)
        self.frame_input.textChanged.connect(self.update_frame_from_input)
        control_layout.addWidget(self.frame_input)

        layout.addLayout(control_layout)

        # Buttons for start/end frame and trimming
        button_layout = QHBoxLayout()

        self.set_start_button = QPushButton("設置開始幀", self)
        self.set_start_button.setEnabled(False)
        self.set_start_button.clicked.connect(self.set_start_frame)
        button_layout.addWidget(self.set_start_button)

        self.set_end_button = QPushButton("設置結束幀", self)
        self.set_end_button.setEnabled(False)
        self.set_end_button.clicked.connect(self.set_end_frame)
        button_layout.addWidget(self.set_end_button)

        self.trim_button = QPushButton("剪輯影片", self)
        self.trim_button.setEnabled(False)
        self.trim_button.clicked.connect(self.trim_video)
        button_layout.addWidget(self.trim_button)

        self.play_pause_button = QPushButton("播放", self)
        self.play_pause_button.setEnabled(False)
        self.play_pause_button.clicked.connect(self.toggle_play_pause)
        button_layout.addWidget(self.play_pause_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)
        self.setAcceptDrops(True)

        # Variables for trimming
        self.start_frame = None
        self.end_frame = None

        # Timer for video playback
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.play_video)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        file_paths = [url.toLocalFile() for url in event.mimeData().urls()]
        if file_paths and file_paths[0].lower().endswith(('.mp4', '.mov', '.avi')):
            self.load_video(file_paths[0])

    def load_video(self, video_path):
        self.video_path = video_path
        self.cap = cv2.VideoCapture(video_path)
        self.total_frames = int(self.cap.get(cv2.CAP_PROP_FRAME_COUNT))
        self.current_frame = 0

        self.slider.setMaximum(self.total_frames - 1)
        self.slider.setEnabled(True)

        self.frame_input.setEnabled(True)
        self.set_start_button.setEnabled(True)
        self.set_end_button.setEnabled(True)
        self.trim_button.setEnabled(True)
        self.play_pause_button.setEnabled(True)

        self.play_video()
        self.timer.start(30)

    def play_video(self):
        if not self.is_playing:
            return  # 如果未播放，直接返回

        if self.cap and self.cap.isOpened():
            self.cap.set(cv2.CAP_PROP_POS_FRAMES, self.current_frame)
            ret, frame = self.cap.read()
            if ret:
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                height, width, channel = frame.shape
                bytes_per_line = 3 * width
                q_image = QImage(frame.data, width, height, bytes_per_line, QImage.Format_RGB888)
                pixmap = QPixmap.fromImage(q_image)
                self.video_label.setPixmap(pixmap.scaled(self.video_label.size(), Qt.KeepAspectRatio))
                self.slider.blockSignals(True)
                self.slider.setValue(self.current_frame)
                self.slider.blockSignals(False)
                self.frame_input.setText(str(self.current_frame))
                self.current_frame += 1
            else:
                self.timer.stop()

    def update_frame_from_slider(self, value):
        self.current_frame = value
        self.cap.set(cv2.CAP_PROP_POS_FRAMES, self.current_frame)
        self.preview_frame()

    def preview_frame(self):
        if self.cap and self.cap.isOpened():
            ret, frame = self.cap.read()
            if ret:
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                height, width, channel = frame.shape
                bytes_per_line = 3 * width
                q_image = QImage(frame.data, width, height, bytes_per_line, QImage.Format_RGB888)
                pixmap = QPixmap.fromImage(q_image)
                self.video_label.setPixmap(pixmap.scaled(self.video_label.size(), Qt.KeepAspectRatio))

    def update_frame_from_input(self):
        try:
            value = int(self.frame_input.text())
            if 0 <= value < self.total_frames:
                self.current_frame = value
                self.cap.set(cv2.CAP_PROP_POS_FRAMES, self.current_frame)
                self.preview_frame()
        except ValueError:
            pass

    def set_start_frame(self):
        self.start_frame = self.current_frame
        print(f"開始幀設置為：{self.start_frame}")

    def set_end_frame(self):
        self.end_frame = self.current_frame
        print(f"結束幀設置為：{self.end_frame}")

    def toggle_play_pause(self):
        if self.is_playing:
            self.timer.stop()
            self.play_pause_button.setText("播放")
        else:
            self.timer.start(30)
            self.play_pause_button.setText("暫停")
        self.is_playing = not self.is_playing

    def trim_video(self):
        if self.start_frame is not None and self.end_frame is not None and self.video_path:
            output_path = os.path.splitext(self.video_path)[0] + f"_trimmed.mp4"
            result = trim_video(self.video_path, self.start_frame, self.end_frame, output_path)
            
            if result:  # 確保剪輯成功
                print(f"影片已剪輯完成！儲存為：{result}")
                open_in_explorer(result)  # 打開檔案總管並聚焦檔案
            else:
                print("影片剪輯失敗！")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VideoTrimWindow()
    window.show()
    sys.exit(app.exec_())
