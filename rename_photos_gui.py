import sys
import os
from collections import Counter
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QLabel, QLineEdit, QFileDialog, 
                             QMessageBox, QHBoxLayout, QFrame)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

class PhotoRenamerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("照片按表对应重命名工具")
        self.resize(550, 480)
        self.setMinimumSize(450, 450)

        # 核心组件
        self.excel_path = ""
        self.folder_path = ""
        self.match_col_input = QLineEdit()
        self.target_col_input = QLineEdit()
        self.lbl_excel = QLabel("尚未选择文件...")
        self.lbl_folder = QLabel("尚未选择文件夹...")

        # 初始化界面
        self.init_ui()

    def init_ui(self):
        # 设置窗口半透明背景 (30% 透明，即 70% 不透明度)
        self.setStyleSheet("""
            QMainWindow { background-color: rgba(26, 43, 76, 180); }
            QWidget { background-color: rgba(26, 43, 76, 180); color: #E0E6ED; font-family: 'PingFang SC'; font-size: 14px; }
            QLabel { font-size: 14px; padding: 2px; }
            QLineEdit { background-color: #FFFFFF; color: #333; border: 1px solid #4A90E2; border-radius: 4px; padding: 6px; font-size: 14px; }
            QPushButton { background-color: #4A90E2; color: white; border: none; border-radius: 6px; padding: 10px; font-size: 15px; font-weight: bold; }
            QPushButton:hover { background-color: #357ABD; }
            QPushButton:pressed { background-color: #1A3B6C; }
        """)

        # 布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(15)

        # 1. Excel 选择区
        layout.addWidget(QLabel("第一步：选择对照表"))
        btn_excel = QPushButton("📂 选择Excel重命名对照表")
        btn_excel.clicked.connect(self.select_excel)
        layout.addWidget(btn_excel)
        self.lbl_excel.setStyleSheet("color: #A0B0C0; font-size: 12px;")
        layout.addWidget(self.lbl_excel)

        # 2. 设置列号
        col_frame = QFrame()
        col_layout = QHBoxLayout(col_frame)
        col_layout.setContentsMargins(0, 0, 0, 0)
        
        col_layout.addWidget(QLabel("【照片原名】列号(如A列填1):"))
        self.match_col_input.setText("1")
        self.match_col_input.setFixedWidth(60)
        col_layout.addWidget(self.match_col_input)
        
        col_layout.addSpacing(20)
        col_layout.addWidget(QLabel("【新文件名】列号(如C列填3):"))
        self.target_col_input.setText("3")
        self.target_col_input.setFixedWidth(60)
        col_layout.addWidget(self.target_col_input)
        layout.addWidget(col_frame)

        # 3. 文件夹选择区
        layout.addWidget(QLabel("第二步：选择照片文件夹"))
        btn_folder = QPushButton("📁 选择照片所在文件夹")
        btn_folder.clicked.connect(self.select_folder)
        layout.addWidget(btn_folder)
        self.lbl_folder.setStyleSheet("color: #A0B0C0; font-size: 12px;")
        layout.addWidget(self.lbl_folder)

        # 弹簧效果，把开始按钮推到底部
        layout.addStretch()

        # 4. 开始按钮
        btn_start = QPushButton("🚀 开始重命名")
        btn_start.setStyleSheet("background-color: #2E7D32; padding: 12px; font-size: 16px;")
        btn_start.clicked.connect(self.start_rename)
        layout.addWidget(btn_start)

    def select_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.excel_path = path
            self.lbl_excel.setText(f"✅ 已选择: {os.path.basename(path)}")
            self.lbl_excel.setStyleSheet("color: #88FF88; font-size: 12px;")

    def select_folder(self):
        path = QFileDialog.getExistingDirectory(self, "选择照片文件夹")
        if path:
            self.folder_path = path
            self.lbl_folder.setText(f"✅ 已选择: {path}")
            self.lbl_folder.setStyleSheet("color: #88FF88; font-size: 12px;")

    def start_rename(self):
        if not self.excel_path or not self.folder_path:
            QMessageBox.warning(self, "提示", "请先选择 Excel 文件和照片文件夹！")
            return

        try:
            c_match = int(self.match_col_input.text().strip()) - 1
            c_target = int(self.target_col_input.text().strip()) - 1
            if c_match < 0 or c_target < 0: raise ValueError()
        except ValueError:
            QMessageBox.critical(self, "错误", "列号必须是大于0的数字！")
            return

        try:
            df = pd.read_excel(self.excel_path, header=None)
            max_cols = len(df.columns)
            if c_match >= max_cols or c_target >= max_cols:
                QMessageBox.critical(self, "错误", f"列号超出范围！表格只有 {max_cols} 列。")
                return

            match_names = df.iloc[:, c_match].astype(str).str.strip()
            target_names = df.iloc[:, c_target].astype(str).str.strip()

            # 重名检测
            name_counts = Counter(match_names)
            duplicate_items = {name for name, count in name_counts.items() if count > 1 and name.lower() != 'nan'}

            # 构建安全映射
            safe_mapping = {}
            for i in range(len(match_names)):
                m_name, t_name = match_names[i], target_names[i]
                if m_name.lower() == 'nan' or t_name.lower() == 'nan': continue
                if m_name not in duplicate_items:
                    safe_mapping[m_name] = t_name

            # 执行重命名
            success_count, fail_count, skip_count = 0, 0, 0
            for filename in os.listdir(self.folder_path):
                file_path = os.path.join(self.folder_path, filename)
                if not os.path.isfile(file_path): continue
                
                pure_name = os.path.splitext(filename)[0].strip()
                _, ext = os.path.splitext(filename)
                
                if pure_name in safe_mapping:
                    new_filename = f"{safe_mapping[pure_name]}{ext}"
                    new_path = os.path.join(self.folder_path, new_filename)
                    if file_path == new_path:
                        skip_count += 1; continue
                    try:
                        os.rename(file_path, new_path)
                        success_count += 1
                    except: fail_count += 1
                elif pure_name.lower() != 'nan':
                    fail_count += 1

            # 结果弹窗
            msg = f"处理完毕！\n✅ 成功重命名: {success_count} 张"
            if skip_count > 0: msg += f"\n⏭️ 跳过(无需修改): {skip_count} 张"
            if fail_count > 0: msg += f"\n❌ 未匹配/失败: {fail_count} 张"
            
            if duplicate_items:
                dup_list = "\n".join(list(duplicate_items)[:10])
                if len(duplicate_items) > 10: dup_list += f"\n...等共 {len(duplicate_items)} 个重名项"
                msg += f"\n\n⚠️ 以下【匹配列内容】存在重名，已严格跳过:\n{dup_list}"
                
            QMessageBox.information(self, "结果", msg)

        except Exception as e:
            QMessageBox.critical(self, "发生错误", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PhotoRenamerApp()
    window.show()
    sys.exit(app.exec_())