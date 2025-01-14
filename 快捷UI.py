import sys
import os
from PyQt5.QtWidgets import QApplication, QLabel, QVBoxLayout, QHBoxLayout, QPushButton, QWidget
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt
import subprocess
import numpy as np
import pickle
import joblib
import json
import pandas as pd
import torch
import cv2
from datetime import datetime


def open_in_explorer(file_path):
    try:
        abs_path = os.path.abspath(file_path)
        # 使用 shell=True 並確保引號正確
        subprocess.run(f'explorer /select,"{abs_path}"', shell=True)
    except Exception as e:
        print(f"無法開啟檔案總管：{e}")


def convert_to_json_serializable(obj):
    if isinstance(obj, dict):
        return {str(k): convert_to_json_serializable(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [convert_to_json_serializable(item) for item in obj]
    elif isinstance(obj, np.ndarray):
        return obj.tolist()
    elif isinstance(obj, pd.DataFrame):
        return obj.to_dict(orient='records')
    elif isinstance(obj, pd.Series):
        return obj.tolist()
    elif isinstance(obj, (np.int32, np.int64, np.float32, np.float64)):
        return obj.item()
    elif isinstance(obj, torch.Tensor):
        return obj.tolist()
    else:
        return obj

def convert_webp_to_jpg(file_path):
    # 檢查副檔名是否為 .webp
    if file_path.lower().endswith(".webp"):
        output_path = os.path.splitext(file_path)[0] + ".jpg"
        try:
            # 使用 ffmpeg 轉換 .webp 為 .jpg
            subprocess.run(["ffmpeg", "-i", file_path, output_path], check=True)
            # 打開檔案總管並聚焦輸出的圖片檔案
            abs_path = os.path.abspath(output_path)
            open_in_explorer(abs_path)
            return f"已轉換為 JPG: {output_path}"
        except subprocess.CalledProcessError as e:
            return f"轉換失敗: {e}"
    return "不是 webp 檔案，無需轉換"

def npy_to_json(npy_file_path):
    json_file_path = os.path.splitext(npy_file_path)[0] + ".json"
    try:
        # 讀取 .npy 檔案
        data = np.load(npy_file_path, allow_pickle=True)
        # 將數據轉換為 JSON 可序列化格式
        serializable_data = convert_to_json_serializable(data)
        # 寫入 JSON 文件
        with open(json_file_path, 'w', encoding='utf-8') as json_file:
            json.dump(serializable_data, json_file, ensure_ascii=False, indent=4)
        # 打開檔案總管並聚焦輸出的 JSON 檔案
        abs_path = os.path.abspath(json_file_path)
        open_in_explorer(abs_path)
        return f"已轉換為 JSON: {json_file_path}"
    except Exception as e:
        # 如果轉換 JSON 失敗，嘗試輸出為 TXT
        txt_file_path = os.path.splitext(npy_file_path)[0] + ".txt"
        data = None  # 先宣告以避免未定義
        try:
            with open(txt_file_path, 'w', encoding='utf-8') as txt_file:
                txt_file.write(str(data))
            # 打開檔案總管並聚焦輸出的 TXT 檔案
            abs_path = os.path.abspath(txt_file_path)
            open_in_explorer(abs_path)
            return f"已轉換為 TXT: {txt_file_path}"
        except Exception as txt_error:
            return f".npy 轉換失敗: {e}，且 TXT 輸出失敗: {txt_error}"

def pkl_to_json(pkl_file_path):
    json_file_path = os.path.splitext(pkl_file_path)[0] + ".json"
    try:
        # 嘗試以 pickle 讀取 .pkl 檔案
        with open(pkl_file_path, 'rb') as f:
            data = pickle.load(f)
        # 將數據轉換為 JSON 可序列化格式
        serializable_data = convert_to_json_serializable(data)
        # 寫入 JSON 文件
        with open(json_file_path, 'w', encoding='utf-8') as json_file:
            json.dump(serializable_data, json_file, ensure_ascii=False, indent=4)
            open_in_explorer(json_file_path)
        return f"已轉換為 JSON: {json_file_path}"
    except pickle.UnpicklingError:
        try:
            # 如果 pickle 失敗，嘗試使用 joblib
            data = joblib.load(pkl_file_path)
            serializable_data = convert_to_json_serializable(data)
            with open(json_file_path, 'w', encoding='utf-8') as json_file:
                json.dump(serializable_data, json_file, ensure_ascii=False, indent=4)
                open_in_explorer(json_file_path)
            return f"已使用 joblib 轉換為 JSON: {json_file_path}"
        except Exception as e:
            return f".pkl 轉換失敗 (joblib): {e}"
    except Exception as e:
        return f".pkl 轉換失敗: {e}"

def compress_pdf(input_path, quality="ebook"):
    """
    使用 Ghostscript 壓縮 PDF 檔案。

    :param input_path: 原始 PDF 檔案路徑
    :param quality: 壓縮品質，可選 screen、ebook、printer、prepress
    :return: 成功訊息或錯誤訊息
    """
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    output_path = f"compressed_{base_name}.pdf"

    # Ghostscript 指令 (Windows 下可能是 gswin64, gswin32c, 或 gs.exe，視安裝情況而定)
    gs_executable = "gswin64"

    gs_command = [
        gs_executable,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS=/{quality}",
        "-dNOPAUSE",
        "-dBATCH",
        f"-sOutputFile={output_path}",
        input_path
    ]

    try:
        subprocess.run(gs_command, check=True)
        return f"壓縮完成！檔案儲存為：{output_path}"
    except FileNotFoundError:
        return ("找不到 Ghostscript 執行檔，請確認系統中是否已安裝，"
                "或將 gs_executable 改為正確的檔名/路徑。")
    except subprocess.CalledProcessError as e:
        return (f"壓縮失敗，Ghostscript 退出代碼：{e.returncode}\n"
                f"請檢查檔案路徑或參數是否正確。")

def folder_to_video(folder_path, frame_rate=30):
    image_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
    if len(image_files) / len(os.listdir(folder_path)) < 0.8:
        return "資料夾內圖片不足 80%，無法生成影片。"

    image_files = sorted(image_files)  # 按名稱排序
    first_image_path = os.path.join(folder_path, image_files[0])
    first_image = cv2.imread(first_image_path)

    if first_image is None:
        return f"無法讀取第一張圖片: {first_image_path}"

    height, width, _ = first_image.shape
    fourcc = cv2.VideoWriter_fourcc(*'mp4v')

    folder_name = os.path.basename(folder_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_video = os.path.join(folder_path, f"{folder_name}_{timestamp}.mp4")

    video_writer = cv2.VideoWriter(output_video, fourcc, frame_rate, (width, height))

    for image_file in image_files:
        img_path = os.path.join(folder_path, image_file)
        img = cv2.imread(img_path)
        if img is None:
            continue
        img = cv2.resize(img, (width, height))
        video_writer.write(img)

    video_writer.release()
    # 打開檔案總管並聚焦輸出的影片檔案
    abs_path = os.path.abspath(output_video)
    open_in_explorer(abs_path)
    return f"影片已輸出為 {output_video}"

def compress_video(input_video):
    video_name, _ = os.path.splitext(os.path.basename(input_video))
    folder_path = os.path.dirname(input_video)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_video = os.path.join(folder_path, f"{video_name}_compressed_{timestamp}.mp4")

    command = [
        "ffmpeg", "-i", input_video,
        "-vcodec", "libx264",
        "-crf", "23",
        "-preset", "medium",
        output_video
    ]

    try:
        subprocess.run(command, check=True)
        # 打開檔案總管並聚焦輸出的影片檔案
        abs_path = os.path.abspath(output_video)
        open_in_explorer(abs_path)
        return f"影片壓縮完成！儲存為：{output_video}"
    except subprocess.CalledProcessError as e:
        return f"影片壓縮失敗！錯誤：{e}"

def handle_drop(event, label):
    event.acceptProposedAction()
    if event.mimeData().hasUrls():
        file_paths = [url.toLocalFile() for url in event.mimeData().urls()]
        if file_paths:
            file_path = file_paths[0]
            if os.path.isdir(file_path):
                result = folder_to_video(file_path)
            else:
                file_name = file_path.split("/")[-1]
                formatted_path = "\n".join([file_path[i:i+50] for i in range(0, len(file_path), 50)])

                conversion_result = ""
                if file_path.lower().endswith(".webp"):
                    conversion_result = convert_webp_to_jpg(file_path)
                elif file_path.lower().endswith(".npy"):
                    conversion_result = npy_to_json(file_path)
                elif file_path.lower().endswith(".pkl"):
                    conversion_result = pkl_to_json(file_path)
                elif file_path.lower().endswith(".pdf"):
                    conversion_result = compress_pdf(file_path)
                elif file_path.lower().endswith(('.mp4', '.mov', '.avi')):
                    conversion_result = compress_video(file_path)
                else:
                    conversion_result = "檔案格式不支援轉換"

                result = f"路徑:\n{formatted_path}\n檔名: {file_name}\n{conversion_result}"

            label.setText(result)
        else:
            label.setText("未檢測到有效的檔案")

class DragDropWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("拖入檔案顯示路徑與檔名")
        self.setGeometry(100, 100, 700, 400)

        self.initUI()

    def initUI(self):
        layout = QHBoxLayout()

        self.label = QLabel("請將檔案拖入此視窗", self)
        self.label.setFont(QFont("Arial", 12))
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("border: 2px dashed #aaa; padding: 20px;")
        self.label.setWordWrap(True)

        left_layout = QVBoxLayout()
        left_layout.addWidget(self.label)

        # Add buttons
        right_layout = QVBoxLayout()

        self.existing_function_button = QPushButton("執行目前功能", self)
        self.existing_function_button.clicked.connect(lambda: self.label.setText("執行目前功能中..."))
        right_layout.addWidget(self.existing_function_button)

        self.video_compress_button = QPushButton("影片壓縮", self)
        self.video_compress_button.clicked.connect(lambda: self.label.setText("請拖入影片檔案進行壓縮..."))
        right_layout.addWidget(self.video_compress_button)

        self.disabled_button = QPushButton("待開發功能", self)
        self.disabled_button.setEnabled(False)
        right_layout.addWidget(self.disabled_button)

        layout.addLayout(left_layout)
        layout.addLayout(right_layout)

        self.setLayout(layout)

        # 啟用拖放
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        handle_drop(event, self.label)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DragDropWindow()
    window.show()
    sys.exit(app.exec_())
