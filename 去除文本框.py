import sys
import os
import traceback
import subprocess
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsmap

# --------------------------
# 核心修复: 强制注册 XML 命名空间
# --------------------------
if 'v' not in nsmap: nsmap['v'] = 'urn:schemas-microsoft-com:vml'
if 'w' not in nsmap: nsmap['w'] = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
if 'o' not in nsmap: nsmap['o'] = 'urn:schemas-microsoft-com:office:office'

# --------------------------
# 依赖检查
# --------------------------
try:
    from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                                 QLabel, QPushButton, QListWidget, QProgressBar, 
                                 QMessageBox, QFileDialog)
    from PyQt6.QtCore import Qt, QThread, pyqtSignal
    from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QFont
except ImportError:
    print("【错误】请安装依赖：pip install python-docx PyQt6")
    sys.exit(1)

# --------------------------
# 核心逻辑 (v10.0 容器级粉碎)
# --------------------------
class Worker(QThread):
    progress_signal = pyqtSignal(int)
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(list)
    error_signal = pyqtSignal(str)

    def __init__(self, file_paths):
        super().__init__()
        self.file_paths = file_paths
        self.is_running = True
        self.generated_files = []

    def run(self):
        total = len(self.file_paths)
        for i, path in enumerate(self.file_paths):
            if not self.is_running: break
            try:
                # 1. 路径处理
                abs_path = os.path.abspath(path)
                dname = os.path.dirname(abs_path)
                fname = os.path.basename(abs_path)
                out_path = os.path.join(dname, f"{os.path.splitext(fname)[0]}_v10重组版.docx")

                self.log_signal.emit(f"正在重组: {fname}")
                
                # 2. 占用检查
                if not self.check_lock(out_path):
                    self.error_signal.emit(f"文件被占用，请关闭：\n{out_path}")
                    return

                # 3. 执行粉碎
                count = self.crush_textboxes(abs_path, out_path)
                
                self.generated_files.append(out_path)
                self.log_signal.emit(f"✔ 完成: {fname} (粉碎 {count} 个容器)")
                
            except Exception as e:
                err = f"{fname} 失败: {str(e)}\n{traceback.format_exc()}"
                print(err)
                self.error_signal.emit(f"严重错误:\n{str(e)}")
            
            self.progress_signal.emit(int((i + 1) / total * 100))
        
        self.finished_signal.emit(self.generated_files)

    def check_lock(self, path):
        if not os.path.exists(path): return True
        try:
            with open(path, 'a'): pass
            return True
        except: return False

    def crush_textboxes(self, input_file, output_file):
        doc = Document(input_file)
        ops_count = 0

        # 获取所有文档区域（正文、表格单元格等）
        # 递归获取所有元素有些复杂，这里主要针对 Body 和 Table Cells
        body_elements = [doc.element.body]
        
        # 扩展：简单的表格支持
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    body_elements.append(cell._element)

        for parent_elm in body_elements:
            # 1. 锁定目标：查找所有容器级节点
            # w:drawing 是新版图形容器
            # w:pict 是旧版 VML 容器 (导致你报错的根源)
            drawings = list(parent_elm.iter(qn('w:drawing')))
            picts = list(parent_elm.iter(qn('w:pict')))
            
            # 合并列表，优先处理深层节点（虽然 docx 是扁平的，但为了安全）
            targets = []
            for d in drawings: targets.append(('drawing', d))
            for p in picts: targets.append(('pict', p))

            # 倒序遍历，确保删除操作安全
            for t_type, node in reversed(targets):
                
                # A. 提取文字
                text_content = self.extract_text(node)
                
                # B. 查找宿主段落 (Anchor Paragraph)
                # 即使 text_content 为空，我们也可能想删除空框
                # 但为了保险，只处理有意义的框或明确的空框
                
                anchor_p = self.find_anchor_paragraph(node)
                
                if text_content:
                    # C. 核心重组逻辑：创建新段落
                    if anchor_p is not None:
                        # 在宿主段落下方插入新段落
                        new_p = self.insert_paragraph_after(anchor_p, text_content)
                        ops_count += 1
                    else:
                        # 极其罕见：找不到宿主，可能是孤立节点，无法安全保留文字位置
                        # 这种情况通常保留原样以免丢失数据，或者打印日志
                        print(f"Skipping orphan node with text: {text_content[:20]}...")
                        continue
                
                # D. 毁灭容器
                # 只有成功提取文字（或者确定是纯空框）后，才删除原节点
                # 这样杜绝了“双重文本”
                self.nuke_node(node)

        doc.save(output_file)
        return ops_count

    def extract_text(self, element):
        """递归提取所有文本节点，并清洗"""
        texts = []
        for t in element.iter(qn('w:t')):
            if t.text:
                texts.append(t.text)
        return "".join(texts).strip()

    def find_anchor_paragraph(self, node):
        """向上追溯找到 w:p 节点"""
        curr = node
        while curr is not None:
            if curr.tag == qn('w:p'):
                return curr
            curr = curr.getparent()
        return None

    def insert_paragraph_after(self, anchor_p, text):
        """在 DOM 树中，在 anchor_p 之后插入一个新的 w:p"""
        # 1. 创建新段落 XML
        new_p = OxmlElement('w:p')
        new_r = OxmlElement('w:r')
        new_t = OxmlElement('w:t')
        
        # 保持空格格式
        if text.startswith(" ") or text.endswith(" "):
            new_t.set(qn('xml:space'), 'preserve')
            
        new_t.text = text
        new_r.append(new_t)
        new_p.append(new_r)
        
        # 2. 插入 DOM
        parent = anchor_p.getparent()
        if parent is not None:
            index = parent.index(anchor_p)
            parent.insert(index + 1, new_p)
        
        return new_p

    def nuke_node(self, node):
        """彻底删除节点"""
        try:
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)
        except:
            pass

# --------------------------
# UI 界面 (保持简洁)
# --------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Word 文本框粉碎机 v10.0 (结构重组版)")
        self.resize(600, 450)
        self.file_queue = []
        self.setAcceptDrops(True) 
        self.init_ui()

    def init_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        
        title = QLabel("🛡️ v10.0 | 容器级粉碎 | 杜绝重复")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color: #C62828; font-weight: 900; font-size: 18px; margin: 15px;")
        layout.addWidget(title)

        self.box = QLabel("将原来的报错文件拖入此处\n(不要拖入 v8/v9 的垃圾文件)")
        self.box.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.box.setStyleSheet("""
            QLabel {
                border: 4px dashed #B71C1C; border-radius: 12px;
                background: #FFEBEE; color: #B71C1C; font-size: 16px; min-height: 120px;
            }
            QLabel:hover { background: #FFCDD2; }
        """)
        self.box.mousePressEvent = self.open_file
        layout.addWidget(self.box)

        self.list = QListWidget()
        layout.addWidget(self.list)
        
        self.prog = QProgressBar()
        layout.addWidget(self.prog)

        btn = QPushButton("立即重组")
        btn.setFixedHeight(55)
        btn.setStyleSheet("""
            QPushButton { background: #B71C1C; color: white; font-size: 18px; font-weight: bold; border-radius: 6px; }
            QPushButton:hover { background: #D32F2F; }
        """)
        btn.clicked.connect(self.start)
        layout.addWidget(btn)
        
        self.stat = QLabel("等待输入...")
        self.stat.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.stat)

    def dragEnterEvent(self, e: QDragEnterEvent):
        if e.mimeData().hasUrls(): e.accept()
        else: e.ignore()

    def dropEvent(self, e: QDropEvent):
        self.add_files(e.mimeData().urls())

    def open_file(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            files, _ = QFileDialog.getOpenFileNames(self, "选文件", "", "Word (*.docx)")
            if files: 
                from PyQt6.QtCore import QUrl
                self.add_files([QUrl.fromLocalFile(f) for f in files])

    def add_files(self, urls):
        n = 0
        for u in urls:
            path = u.toLocalFile()
            if sys.platform == 'win32' and path.startswith("/"): path = path[1:]
            if path.endswith(".docx") and not os.path.basename(path).startswith("~$"):
                path = os.path.normpath(path)
                if path not in self.file_queue:
                    self.file_queue.append(path)
                    self.list.addItem(path)
                    n += 1
        if n: self.stat.setText(f"已就绪 {len(self.file_queue)} 个文件")

    def start(self):
        if not self.file_queue: return QMessageBox.warning(self, "!", "先拖入文件")
        self.worker = Worker(self.file_queue)
        self.worker.progress_signal.connect(self.prog.setValue)
        self.worker.log_signal.connect(self.stat.setText)
        self.worker.error_signal.connect(lambda m: QMessageBox.critical(self, "Error", m))
        self.worker.finished_signal.connect(self.done)
        self.worker.start()

    def done(self, files):
        self.stat.setText("全部完成")
        self.file_queue.clear()
        self.list.clear()
        if files and QMessageBox.question(self, "完成", "打开文件夹？", 
           QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            d = os.path.dirname(files[0])
            if sys.platform == 'win32': subprocess.Popen(['explorer', d])
            else: subprocess.Popen(['open', d])

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei", 9))
    win = MainWindow()
    win.show()
    sys.exit(app.exec())