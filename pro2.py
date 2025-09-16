import sys
import os
import json
import warnings

# 屏蔽 openpyxl 的默认样式警告，避免不必要的控制台输出
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTextEdit, QLabel, QFileDialog, QMessageBox,
    QListWidget, QListWidgetItem, QLineEdit, QComboBox, QInputDialog,
    QDialog, QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView, QSplitter, QCheckBox
)
from PyQt6.QtCore import QDateTime, Qt, QUrl
from PyQt6.QtGui import QDesktopServices, QDragEnterEvent, QDropEvent


# --- 新的字段配置对话框，支持多可配置字段列配置 ---
class OutputConfigDialog(QDialog):
    def __init__(self, initial_general_map, initial_value_map, parent=None):
        super().__init__(parent)
        self.setWindowTitle("配置输出字段和顺序")
        self.resize(800, 600)

        main_layout = QVBoxLayout(self)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        # 左侧：通用字段配置
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.addWidget(QLabel("通用字段（序号、索引、转换后列）："))
        self.general_table = QTableWidget(0, 2)
        self.general_table.setHorizontalHeaderLabels(["原始字段", "新名称"])
        self.general_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.general_table.cellDoubleClicked.connect(self.edit_general_cell)
        self.populate_general_table(initial_general_map)
        left_layout.addWidget(self.general_table)

        general_btn_layout = QHBoxLayout()
        self.btn_general_up = QPushButton("上移")
        self.btn_general_up.clicked.connect(lambda: self.move_item(self.general_table, -1))
        self.btn_general_down = QPushButton("下移")
        self.btn_general_down.clicked.connect(lambda: self.move_item(self.general_table, 1))
        general_btn_layout.addWidget(self.btn_general_up)
        general_btn_layout.addWidget(self.btn_general_down)
        left_layout.addLayout(general_btn_layout)

        # 右侧：可配置字段列配置
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.addWidget(QLabel("可配置的字段（双击重命名）："))
        self.value_table = QTableWidget(0, 2)
        self.value_table.setHorizontalHeaderLabels(["原始文件名", "新名称"])
        self.value_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.value_table.cellDoubleClicked.connect(self.edit_value_cell)
        self.populate_value_table(initial_value_map)
        right_layout.addWidget(self.value_table)

        # 添加到splitter
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        main_layout.addWidget(splitter)

        # 底部按钮
        ok_btn = QPushButton("确定")
        ok_btn.clicked.connect(self.accept)
        main_layout.addWidget(ok_btn, alignment=Qt.AlignmentFlag.AlignRight)

        self.general_result = {}
        self.value_result = {}

    def populate_general_table(self, fields_dict):
        self.general_table.setRowCount(len(fields_dict))
        for i, (original, new) in enumerate(fields_dict.items()):
            original_item = QTableWidgetItem(original)
            original_item.setFlags(original_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.general_table.setItem(i, 0, original_item)
            self.general_table.setItem(i, 1, QTableWidgetItem(new))

    def populate_value_table(self, fields_dict):
        self.value_table.setRowCount(len(fields_dict))
        for i, (original, new) in enumerate(fields_dict.items()):
            original_item = QTableWidgetItem(original)
            original_item.setFlags(original_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.value_table.setItem(i, 0, original_item)
            self.value_table.setItem(i, 1, QTableWidgetItem(new))

    def edit_general_cell(self, row, column):
        if column == 1:
            item = self.general_table.item(row, column)
            new_text, ok = QInputDialog.getText(self, "重命名", "请输入新的名称：", text=item.text())
            if ok and new_text:
                item.setText(new_text)

    def edit_value_cell(self, row, column):
        if column == 1:
            item = self.value_table.item(row, column)
            new_text, ok = QInputDialog.getText(self, "重命名", "请输入新的名称：", text=item.text())
            if ok and new_text:
                item.setText(new_text)

    def move_item(self, table, direction):
        current_row = table.currentRow()
        if direction == -1 and current_row > 0:
            table.insertRow(current_row - 1)
            for col in range(table.columnCount()):
                item = table.takeItem(current_row + 1, col)
                table.setItem(current_row - 1, col, item)
            table.removeRow(current_row + 1)
            table.setCurrentCell(current_row - 1, 0)
        elif direction == 1 and current_row < table.rowCount() - 1:
            table.insertRow(current_row + 2)
            for col in range(table.columnCount()):
                item = table.takeItem(current_row, col)
                table.setItem(current_row + 2, col, item)
            table.removeRow(current_row)
            table.setCurrentCell(current_row + 1, 0)

    def accept(self):
        self.general_result = {}
        for row in range(self.general_table.rowCount()):
            original = self.general_table.item(row, 0).text()
            new = self.general_table.item(row, 1).text()
            self.general_result[original] = new

        self.value_result = {}
        for row in range(self.value_table.rowCount()):
            original = self.value_table.item(row, 0).text()
            new = self.value_table.item(row, 1).text()
            self.value_result[original] = new

        super().accept()


class ExcelCleanerGeneral(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("宽表转长表通用工具")
        self.resize(1000, 720)

        self.input_files = []
        self.output_files = []
        self.df_cache = None
        self.current_columns = []
        self.export_folder = None
        self.input_folder = None

        self.config_confirmed = False

        self.general_output_map = {}
        self.value_output_map = {}

        self.rule = {
            "selected_columns": [],
            "index_column": None,
            "index_alias": None,
            "value_column_alias": "日期",
            "output_name_template": "清洗_{basename}.xlsx",
            "expand_mode": "index_then_value",
            "enable_serial_number": True,
            "enable_trim_and_prefix": True,
            "data_prefix": "#",
            "general_output_map": {}  # 新增的规则字段
        }

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)

        # 第一行：导入 / 选择导出目录 / 导出名模板 / 修改所选导出名
        row1 = QHBoxLayout()
        self.btn_import = QPushButton("【1】导入Excel文件（可批量拖拽）")
        self.btn_import.clicked.connect(self.import_files)
        row1.addWidget(self.btn_import)

        self.btn_select_export_folder = QPushButton("【2】选择导出文件夹")
        self.btn_select_export_folder.clicked.connect(self.select_export_folder)
        row1.addWidget(self.btn_select_export_folder)

        self.btn_open_input_folder = QPushButton("打开导入文件夹")
        self.btn_open_input_folder.clicked.connect(self.open_input_folder)
        row1.addWidget(self.btn_open_input_folder)

        row1.addWidget(QLabel("导出文件名模板："))
        self.edit_export_name = QLineEdit()
        self.edit_export_name.setPlaceholderText("支持 {basename} 占位，例如：清洗_{basename}.xlsx")
        self.edit_export_name.setText(self.rule["output_name_template"])
        row1.addWidget(self.edit_export_name)

        self.btn_edit_selected_output = QPushButton("修改选中文件导出名")
        self.btn_edit_selected_output.clicked.connect(self.edit_selected_output_names)
        row1.addWidget(self.btn_edit_selected_output)

        main_layout.addLayout(row1)

        # 第二行：列选择区 + 控制区 + 注意事项区
        row2 = QHBoxLayout()

        # 列选择区
        col_box_layout = QVBoxLayout()
        col_box_layout.addWidget(QLabel("可选列（勾选后参与展开，默认全选；双击切换勾选）："))
        self.list_columns = QListWidget()
        self.list_columns.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.list_columns.itemDoubleClicked.connect(self.toggle_column_selection)
        col_box_layout.addWidget(self.list_columns)

        select_btns_layout = QHBoxLayout()
        self.btn_select_all = QPushButton("全选")
        self.btn_select_all.clicked.connect(self.select_all_columns)
        select_btns_layout.addWidget(self.btn_select_all)

        self.btn_deselect_all = QPushButton("全不选")
        self.btn_deselect_all.clicked.connect(self.deselect_all_columns)
        select_btns_layout.addWidget(self.btn_deselect_all)

        col_box_layout.addLayout(select_btns_layout)
        row2.addLayout(col_box_layout, 2)

        # 控制区
        ctrl_layout = QVBoxLayout()
        ctrl_layout.addWidget(QLabel("选择索引列（id 列）："))
        self.combo_index = QComboBox()
        ctrl_layout.addWidget(self.combo_index)

        ctrl_layout.addWidget(QLabel("索引列输出别名（例如：科室）："))
        self.edit_index_alias = QLineEdit()
        ctrl_layout.addWidget(self.edit_index_alias)

        ctrl_layout.addWidget(QLabel("展开后列名（value 列名，例如：日期）："))
        self.edit_value_alias = QLineEdit()
        self.edit_value_alias.setText("日期")
        ctrl_layout.addWidget(self.edit_value_alias)

        ctrl_layout.addWidget(QLabel("展开顺序："))
        self.combo_expand_mode = QComboBox()
        self.combo_expand_mode.addItems(["按索引先展开（每个索引展开所有 value）", "按列先展开（每列展开所有索引）"])
        ctrl_layout.addWidget(self.combo_expand_mode)

        # 增加序号列功能（默认开启）
        self.cb_add_index_column = QCheckBox("增加序号列")
        self.cb_add_index_column.setChecked(self.rule["enable_serial_number"])
        self.cb_add_index_column.stateChanged.connect(self.update_rule)
        ctrl_layout.addWidget(self.cb_add_index_column)

        # 新增数据清理和前缀选项
        clean_layout = QHBoxLayout()
        self.cb_trim_and_prefix = QCheckBox("启用数据清理和添加前缀")
        self.cb_trim_and_prefix.setChecked(self.rule["enable_trim_and_prefix"])
        self.cb_trim_and_prefix.stateChanged.connect(self.update_rule)
        clean_layout.addWidget(self.cb_trim_and_prefix)
        self.edit_data_prefix = QLineEdit()
        self.edit_data_prefix.setText(self.rule["data_prefix"])
        self.edit_data_prefix.textChanged.connect(self.update_rule)
        clean_layout.addWidget(QLabel("前缀："))
        clean_layout.addWidget(self.edit_data_prefix)
        ctrl_layout.addLayout(clean_layout)

        self.btn_configure_output = QPushButton("【3】配置输出字段和顺序")
        self.btn_configure_output.clicked.connect(self.configure_output_fields)
        ctrl_layout.addWidget(self.btn_configure_output)

        rules_layout = QHBoxLayout()
        self.btn_save_rule = QPushButton("保存规则 (JSON)")
        self.btn_save_rule.clicked.connect(self.save_rule)
        self.btn_load_rule = QPushButton("加载规则 (JSON)")
        self.btn_load_rule.clicked.connect(self.load_rule)
        rules_layout.addWidget(self.btn_save_rule)
        rules_layout.addWidget(self.btn_load_rule)
        ctrl_layout.addLayout(rules_layout)

        row2.addLayout(ctrl_layout, 1)

        # 注意事项区
        tips_layout = QVBoxLayout()
        tips_layout.addWidget(QLabel("操作注意事项："))
        self.tips_text = QTextEdit()
        self.tips_text.setReadOnly(True)
        self.tips_text.setMarkdown(
            "**1. 拖拽导入：**\n"
            "将 .xlsx 或 .xls 文件拖拽到程序窗口，快速批量导入。\n\n"
            "**2. 表头一致：**\n"
            "批量处理时，请确保所有文件的表头（第一行标题）完全一致，否则程序会报错并终止。\n\n"
            "**3. 索引列：**\n"
            "选择一个唯一标识每行记录的列作为索引列（ID），例如工号或姓名。\n\n"
            "**4. 配置输出：**\n"
            "转换前务必点击【3】配置输出，以确保最终表格结构符合要求。**此为强制步骤！**\n\n"
            "**5. 规则功能：**\n"
            "你可以将当前所有配置（包括列选择、别名和通用字段顺序）保存为规则文件，以便下次直接加载使用。\n\n"
            "**6. 数据清理：**\n"
            "勾选“启用数据清理和添加前缀”可自动去除单元格前后空格。你也可以自定义前缀，例如 #。\n\n"
            "**7. 批量导出：**\n"
            "程序会根据你的文件名模板，依次处理所有导入文件，并导出到指定文件夹。"
        )
        tips_layout.addWidget(self.tips_text)
        row2.addLayout(tips_layout, 1)

        main_layout.addLayout(row2)

        main_layout.addWidget(QLabel("导入的文件（双击项以从列表删除）："))
        self.file_list_widget = QListWidget()
        self.file_list_widget.itemDoubleClicked.connect(self.remove_input_file)
        main_layout.addWidget(self.file_list_widget)

        btn_row = QHBoxLayout()
        self.btn_convert = QPushButton("【4】开始转换并导出（批量）")
        self.btn_convert.clicked.connect(self.convert_and_export_all)
        btn_row.addWidget(self.btn_convert)

        self.btn_export_single = QPushButton("测试_导出单个文件（仅首个）")
        self.btn_export_single.clicked.connect(self.export_current_single)
        btn_row.addWidget(self.btn_export_single)

        self.btn_clear = QPushButton("初始化")
        self.btn_clear.clicked.connect(self.initialize_app)
        btn_row.addWidget(self.btn_clear)

        self.btn_open_export_folder = QPushButton("打开导出文件夹")
        self.btn_open_export_folder.clicked.connect(self.open_export_folder)
        btn_row.addWidget(self.btn_open_export_folder)

        main_layout.addLayout(btn_row)

        main_layout.addWidget(QLabel("运行日志："))
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        main_layout.addWidget(self.log_text, 2)

        self.footer_label = QLabel("开发者: 欧星星 | 宽表转长表通用工具 v2.0")
        self.footer_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.footer_label.setStyleSheet("color: gray; font-size: 12px;")
        main_layout.addWidget(self.footer_label)

        self.setAcceptDrops(True)

    def log(self, message: str, error: bool = False):
        ts = QDateTime.currentDateTime().toString("yyyy-MM-dd hh:mm:ss")
        line = f"[{ts}] {message}"
        if error:
            line = f"<span style='color:red;'>{line}</span>"
        else:
            line = f"<span>{line}</span>"
        self.log_text.append(line)
        self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())

    def update_rule(self):
        self.rule["enable_trim_and_prefix"] = self.cb_trim_and_prefix.isChecked()
        self.rule["data_prefix"] = self.edit_data_prefix.text()
        self.rule["enable_serial_number"] = self.cb_add_index_column.isChecked()
        self.log("数据清理设置已更新到当前规则。")
        self.config_confirmed = False

    def initialize_app(self):
        confirm = QMessageBox.question(self, "确认初始化",
                                       "确定要清空所有记录、规则和日志，回到初始状态吗？此操作不可撤销。",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self.input_files.clear()
        self.output_files.clear()
        self.df_cache = None
        self.current_columns = []
        self.export_folder = None
        self.input_folder = None
        self.file_list_widget.clear()
        self.list_columns.clear()
        self.combo_index.clear()
        self.cb_trim_and_prefix.setChecked(True)
        self.edit_data_prefix.setText("#")
        self.config_confirmed = False

        self.rule = {
            "selected_columns": [],
            "index_column": None,
            "index_alias": None,
            "value_column_alias": "日期",
            "output_name_template": "清洗_{basename}.xlsx",
            "expand_mode": "index_then_value",
            "enable_serial_number": True,
            "enable_trim_and_prefix": True,
            "data_prefix": "#",
            "general_output_map": {}
        }
        self.edit_export_name.setText(self.rule["output_name_template"])
        self.edit_index_alias.clear()
        self.edit_value_alias.setText("日期")
        self.cb_add_index_column.setChecked(self.rule["enable_serial_number"])

        self.log_text.clear()
        self.log("程序已初始化，所有记录和设置均已清空。")
        self.general_output_map = {}
        self.value_output_map = {}

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        excel_files = [f for f in files if f.lower().endswith((".xls", ".xlsx"))]
        if not excel_files:
            self.log("没有找到 Excel 文件（拖拽忽略）。")
            return
        for f in excel_files:
            self.add_input_file(f)

    def import_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择 Excel 文件", "", "Excel Files (*.xlsx *.xls)")
        for f in files:
            self.add_input_file(f)

    def add_input_file(self, path):
        if not os.path.exists(path):
            self.log(f"文件不存在：{path}", error=True)
            return
        if path in self.input_files:
            self.log(f"已存在文件：{os.path.basename(path)}（跳过）")
            return

        self.input_files.append(path)
        self.input_folder = os.path.dirname(path)

        template = self.edit_export_name.text().strip() or self.rule.get("output_name_template", "清洗_{basename}.xlsx")
        basename = os.path.splitext(os.path.basename(path))[0]
        default_out = template.format(basename=basename)
        if not default_out.lower().endswith(".xlsx"):
            default_out += ".xlsx"
        self.output_files.append(default_out)
        self.file_list_widget.addItem(f"{os.path.basename(path)} -> {default_out}")
        self.log(f"已导入文件：{os.path.basename(path)}")
        self.config_confirmed = False

        file_basename = self.choose_basename_for_file(path)
        if file_basename not in self.value_output_map:
            self.value_output_map[file_basename] = file_basename

        if len(self.input_files) == 1:
            try:
                engine = 'xlrd' if path.lower().endswith('.xls') else 'openpyxl'
                df = pd.read_excel(path, engine=engine)
                self.df_cache = df
                self.current_columns = [str(c).strip() for c in df.columns]
                self.populate_column_ui(selected_cols=self.rule.get("selected_columns"))

                if self.rule.get("index_column") and self.rule.get("index_column") in self.current_columns:
                    self.combo_index.setCurrentText(self.rule.get("index_column"))
                elif self.current_columns:
                    self.combo_index.setCurrentIndex(0)
            except Exception as e:
                self.log(f"读取表头失败：{e}", error=True)
                self.input_files.pop()
                self.output_files.pop()
                self.file_list_widget.takeItem(self.file_list_widget.count() - 1)

    def remove_input_file(self, item: QListWidgetItem):
        row = self.file_list_widget.row(item)
        if row < 0 or row >= len(self.input_files):
            return
        removed_path = self.input_files.pop(row)
        removed_out = self.output_files.pop(row)
        self.file_list_widget.takeItem(row)
        self.config_confirmed = False

        removed_basename = self.choose_basename_for_file(removed_path)
        if removed_basename in self.value_output_map:
            del self.value_output_map[removed_basename]

        self.log(f"已从导入列表移除：{os.path.basename(removed_path)} -> {removed_out}")
        if not self.input_files:
            self.df_cache = None
            self.current_columns = []
            self.list_columns.clear()
            self.combo_index.clear()
            self.input_folder = None
            self.general_output_map = {}
            self.value_output_map = {}

    def open_input_folder(self):
        if not self.input_folder:
            QMessageBox.warning(self, "提示", "请先导入文件！")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(self.input_folder))

    def select_export_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择导出文件夹")
        if not folder:
            return
        self.export_folder = folder
        self.log(f"已选择导出文件夹：{folder}")

    def open_export_folder(self):
        if not self.export_folder:
            QMessageBox.warning(self, "提示", "请先选择导出文件夹")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(self.export_folder))

    def toggle_column_selection(self, item: QListWidgetItem):
        if item.checkState() == Qt.CheckState.Checked:
            item.setCheckState(Qt.CheckState.Unchecked)
        else:
            item.setCheckState(Qt.CheckState.Checked)
        self.config_confirmed = False

    def populate_column_ui(self, selected_cols=None):
        self.list_columns.clear()
        self.combo_index.clear()
        if not self.current_columns:
            return

        for col in self.current_columns:
            it = QListWidgetItem(col)
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            if selected_cols is not None and col in selected_cols:
                it.setCheckState(Qt.CheckState.Checked)
            else:
                it.setCheckState(Qt.CheckState.Unchecked)
            self.list_columns.addItem(it)
            self.combo_index.addItem(col)

        if not selected_cols:
            self.select_all_columns()

        if self.current_columns:
            self.combo_index.setCurrentIndex(0)
        self.config_confirmed = False

    def select_all_columns(self):
        for i in range(self.list_columns.count()):
            item = self.list_columns.item(i)
            item.setCheckState(Qt.CheckState.Checked)
        self.log("已全选所有列")
        self.config_confirmed = False

    def deselect_all_columns(self):
        for i in range(self.list_columns.count()):
            item = self.list_columns.item(i)
            item.setCheckState(Qt.CheckState.Unchecked)
        self.log("已取消全选所有列")
        self.config_confirmed = False

    def edit_selected_output_names(self):
        sel_items = self.file_list_widget.selectedItems()
        if not sel_items:
            QMessageBox.warning(self, "提示", "请先选中要修改的文件项（支持多选）")
            return
        for item in sel_items:
            row = self.file_list_widget.row(item)
            current_out = self.output_files[row]
            new_name, ok = QInputDialog.getText(self, "修改导出文件名", "请输入新的导出文件名（含 .xlsx）：",
                                                text=current_out)
            if ok and new_name:
                if not new_name.lower().endswith(".xlsx"):
                    new_name = new_name + ".xlsx"
                self.output_files[row] = new_name
                basename = os.path.basename(self.input_files[row])
                item.setText(f"{basename} -> {new_name}")
                self.log(f"已修改导出名：{basename} -> {new_name}")
        self.config_confirmed = False

    def build_rule_from_ui(self):
        selected_columns = []
        for i in range(self.list_columns.count()):
            item = self.list_columns.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                selected_columns.append(item.text())
        index_col = self.combo_index.currentText() if self.combo_index.count() > 0 else None

        rule = {
            "selected_columns": selected_columns,
            "index_column": index_col,
            "index_alias": self.edit_index_alias.text().strip() or index_col,
            "value_column_alias": self.edit_value_alias.text().strip() or "日期",
            "output_name_template": self.edit_export_name.text().strip() or self.rule.get("output_name_template"),
            "expand_mode": "index_then_value" if self.combo_expand_mode.currentIndex() == 0 else "value_then_index",
            "enable_serial_number": self.cb_add_index_column.isChecked(),
            "enable_trim_and_prefix": self.cb_trim_and_prefix.isChecked(),
            "data_prefix": self.edit_data_prefix.text(),
            "general_output_map": self.general_output_map
        }
        return rule

    def configure_output_fields(self):
        if not self.input_files:
            QMessageBox.warning(self, "提示", "请先导入文件。")
            return

        current_rule = self.build_rule_from_ui()
        index_alias = current_rule.get("index_alias") or current_rule.get("index_column")
        value_alias = current_rule.get("value_column_alias")
        enable_serial = current_rule.get("enable_serial_number")

        initial_general_map = current_rule["general_output_map"].copy()
        if not initial_general_map:
            if enable_serial:
                initial_general_map["序号"] = "序号"
            initial_general_map["索引列名"] = index_alias
            initial_general_map["转换后列名"] = value_alias
        else:
            if enable_serial and "序号" not in initial_general_map:
                initial_general_map = {"序号": "序号", **initial_general_map}
            elif not enable_serial and "序号" in initial_general_map:
                del initial_general_map["序号"]
            if "索引列名" not in initial_general_map or initial_general_map["索引列名"] != index_alias:
                initial_general_map["索引列名"] = index_alias
            if "转换后列名" not in initial_general_map or initial_general_map["转换后列名"] != value_alias:
                initial_general_map["转换后列名"] = value_alias

        dialog = OutputConfigDialog(initial_general_map, self.value_output_map, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.general_output_map = dialog.general_result
            self.value_output_map = dialog.value_result
            self.log("输出字段配置已更新。")
            self.config_confirmed = True

    def apply_rule_to_ui(self, rule: dict):
        if not self.input_files:
            QMessageBox.warning(self, "提示", "请先导入文件以加载表头，再加载规则。")
            self.log("加载规则失败：没有导入文件。", error=True)
            return

        current_cols_set = set(self.current_columns)
        rule_cols_set = set(rule.get("selected_columns", []))
        index_col = rule.get("index_column")

        if not index_col or index_col not in current_cols_set:
            self.log(f"规则中的索引列 '{index_col}' 不存在于当前文件中，规则不适用。", error=True)
            QMessageBox.warning(self, "规则不适用",
                                f"规则中的索引列 '{index_col}' 不存在于当前文件中。\n请检查文件结构或手动配置。")
            return

        missing_cols = rule_cols_set - current_cols_set
        if missing_cols:
            self.log(f"规则中的部分列在当前文件中不存在：{', '.join(missing_cols)}，规则不适用。", error=True)
            QMessageBox.warning(self, "规则不适用",
                                f"规则中的部分列在当前文件中不存在：\n{', '.join(missing_cols)}\n该规则不完全适用于当前文件，请检查文件结构或手动配置。")
            return

        self.rule = rule

        self.populate_column_ui(selected_cols=rule_cols_set)

        self.combo_index.setCurrentText(index_col)
        self.edit_index_alias.setText(rule.get("index_alias", ""))
        self.edit_value_alias.setText(rule.get("value_column_alias", "日期"))
        self.edit_export_name.setText(rule.get("output_name_template", "清洗_{basename}.xlsx"))
        self.combo_expand_mode.setCurrentIndex(0 if rule.get("expand_mode") == "index_then_value" else 1)

        self.cb_add_index_column.setChecked(rule.get("enable_serial_number", True))
        self.cb_trim_and_prefix.setChecked(rule.get("enable_trim_and_prefix", True))
        self.edit_data_prefix.setText(rule.get("data_prefix", "#"))

        self.general_output_map = rule.get("general_output_map", {})
        self.value_output_map = {self.choose_basename_for_file(f): self.choose_basename_for_file(f) for f in
                                 self.input_files}

        self.config_confirmed = False

        self.log("已成功加载并应用规则。请务必点击【3】配置输出字段和顺序进行确认。")

    def save_rule(self):
        if not self.general_output_map:
            QMessageBox.warning(self, "提示", "请先点击【3】配置输出字段并确认，否则通用字段不会被保存。")
            return

        rule = self.build_rule_from_ui()

        if not rule["index_column"] or not rule["selected_columns"]:
            QMessageBox.warning(self, "提示", "请先选择索引列和至少一个要展开的列。")
            return

        # 确保只将 general_output_map 写入规则
        rule_to_save = rule.copy()
        rule_to_save["general_output_map"] = self.general_output_map
        if "value_output_map" in rule_to_save:
            del rule_to_save["value_output_map"]

        path, _ = QFileDialog.getSaveFileName(self, "保存规则为 JSON", "", "JSON 文件 (*.json)")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(rule_to_save, f, ensure_ascii=False, indent=2)
            self.log(f"规则已保存到: {path}")
        except Exception as e:
            self.log(f"保存规则失败: {e}", error=True)

    def load_rule(self):
        if not self.input_files:
            QMessageBox.warning(self, "提示", "请先导入文件以加载表头，再加载规则。")
            return
        path, _ = QFileDialog.getOpenFileName(self, "加载规则 (JSON)", "", "JSON 文件 (*.json)")
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                rule = json.load(f)
            self.apply_rule_to_ui(rule)
        except Exception as e:
            self.log(f"加载规则失败: {e}", error=True)

    def convert_one_df(self, df: pd.DataFrame, rule: dict, metric_name: str):
        id_col = rule["index_column"]
        selected_cols_from_rule = rule["selected_columns"]

        if id_col not in df.columns:
            raise ValueError(f"索引列 '{id_col}' 不存在于当前文件中。")
        df.columns = [str(c).strip() for c in df.columns]

        value_cols = [c for c in selected_cols_from_rule if c in df.columns and c != id_col]
        if not value_cols:
            raise ValueError("没有可用的列进行展开。")

        value_name = rule["value_column_alias"]
        index_alias = rule["index_alias"] or id_col

        melted = pd.melt(df, id_vars=[id_col], value_vars=value_cols,
                         var_name=value_name, value_name=metric_name)

        if rule["expand_mode"] == "index_then_value":
            index_order_map = {v: i for i, v in enumerate(df[id_col].tolist())}
            value_order_map = {v: i for i, v in enumerate(value_cols)}
            melted["_idx_order"] = melted[id_col].map(index_order_map)
            melted["_val_order"] = melted[value_name].map(value_order_map)
            melted = melted.sort_values(by=["_idx_order", "_val_order"]).reset_index(drop=True)
            melted.drop(columns=["_idx_order", "_val_order"], inplace=True)
        else:
            value_order_map = {v: i for i, v in enumerate(value_cols)}
            index_order_map = {v: i for i, v in enumerate(df[id_col].tolist())}
            melted["_val_order"] = melted[value_name].map(value_order_map)
            melted["_idx_order"] = melted[id_col].map(index_order_map)
            melted = melted.sort_values(by=["_val_order", "_idx_order"]).reset_index(drop=True)
            melted.drop(columns=["_idx_order", "_idx_order"], inplace=True)

        melted = melted.rename(columns={id_col: index_alias})

        if rule.get("enable_trim_and_prefix", True):
            data_prefix = rule.get("data_prefix", "#")
            target_cols = [index_alias, metric_name]
            for col in target_cols:
                if col in melted.columns:
                    melted[col] = melted[col].astype(str).apply(
                        lambda x: (data_prefix * (len(x) - len(x.lstrip()))) + x.lstrip()
                    )

        # 优先使用规则中保存的 general_output_map
        if rule.get("general_output_map"):
            output_col_map = rule["general_output_map"].copy()
        else:
            output_col_map = self.general_output_map.copy()

        value_col_new_name = self.value_output_map.get(os.path.splitext(metric_name)[0], metric_name)

        # 新增“可配置字段”这一列的映射
        output_col_map["可配置字段"] = value_col_new_name

        output_df = pd.DataFrame()
        final_columns = list(output_col_map.values())
        output_df = pd.DataFrame(columns=final_columns)

        for original_name, new_name in output_col_map.items():
            if original_name == "序号":
                if rule.get("enable_serial_number"):
                    output_df[new_name] = range(1, len(melted) + 1)
            elif original_name == "索引列名":
                output_df[new_name] = melted[index_alias]
            elif original_name == "转换后列名":
                output_df[new_name] = melted[value_name]
            elif original_name == "可配置字段":
                output_df[new_name] = melted[metric_name]
            else:
                if original_name in melted.columns:
                    output_df[new_name] = melted[original_name]
                else:
                    output_df[new_name] = pd.NA
        return output_df

    def choose_basename_for_file(self, file_path):
        return os.path.splitext(os.path.basename(file_path))[0]

    def ensure_xlsx_ext(self, name: str):
        return name if name.lower().endswith(".xlsx") else name + ".xlsx"

    def convert_and_export_all(self):
        if not self.input_files:
            QMessageBox.warning(self, "提示", "请先导入文件")
            return
        if not self.export_folder:
            QMessageBox.warning(self, "提示", "请先选择导出文件夹")
            return

        if not self.config_confirmed:
            QMessageBox.warning(self, "警告", "请先点击【3】配置输出字段和顺序按钮进行配置确认。")
            self.log("操作失败：请先配置输出字段和顺序。", error=True)
            return

        rule = self.build_rule_from_ui()
        if not rule["index_column"] or not rule["selected_columns"]:
            QMessageBox.warning(self, "提示", "请先选择索引列和至少一个要展开的列")
            return

        if not self.general_output_map:
            QMessageBox.warning(self, "提示", "请先点击【3】配置输出字段和顺序”按钮进行配置。")
            return

        base_columns = None
        for path in self.input_files:
            try:
                engine = 'xlrd' if path.lower().endswith('.xls') else 'openpyxl'
                df = pd.read_excel(path, engine=engine)
                current_columns = [str(c).strip() for c in df.columns]
                if base_columns is None:
                    base_columns = current_columns
                elif set(base_columns) != set(current_columns):
                    QMessageBox.critical(self, "错误",
                                         f"检测到 **{os.path.basename(path)}** 等文件表头结构不一致，请确保批量处理的所有文件的表头字段完全相同。")
                    self.log(f"文件 {os.path.basename(path)} 表头不一致，批量处理失败。", error=True)
                    return
            except Exception as e:
                QMessageBox.critical(self, "错误", f"读取文件 **{os.path.basename(path)}** 表头失败：{e}")
                self.log(f"读取文件 {os.path.basename(path)} 失败，批量处理终止。", error=True)
                return

        for idx, path in enumerate(list(self.input_files)):
            try:
                self.log(f"开始处理：{os.path.basename(path)}")
                engine = 'xlrd' if path.lower().endswith('.xls') else 'openpyxl'
                df = pd.read_excel(path, engine=engine)
                metric_name = self.choose_basename_for_file(path)
                out_df = self.convert_one_df(df, rule, metric_name)
                out_name = self.output_files[idx] if idx < len(self.output_files) and self.output_files[idx] else \
                    self.edit_export_name.text().strip().format(basename=os.path.splitext(os.path.basename(path))[0])
                out_name = self.ensure_xlsx_ext(out_name)
                out_path = os.path.join(self.export_folder, out_name)
                out_df.to_excel(out_path, index=False, engine="openpyxl")
                self.log(f"成功导出：{out_name} （{len(out_df)} 行）")

            except Exception as e:
                self.log(f"处理文件出错：{path} 错误：{e}", error=True)

    def export_current_single(self):
        if not self.input_files:
            QMessageBox.warning(self, "提示", "请先导入文件")
            return
        if not self.export_folder:
            QMessageBox.warning(self, "提示", "请先选择导出文件夹")
            return

        if not self.config_confirmed:
            QMessageBox.warning(self, "警告", "请先点击【3】配置输出字段和顺序按钮进行配置确认。")
            self.log("操作失败：请先配置输出字段和顺序。", error=True)
            return

        path = self.input_files[0]
        rule = self.build_rule_from_ui()

        if not self.general_output_map:
            QMessageBox.warning(self, "提示", "请先点击【3】配置输出字段和顺序”按钮进行配置。")
            return

        try:
            engine = 'xlrd' if path.lower().endswith('.xls') else 'openpyxl'
            df = pd.read_excel(path, engine=engine)
            metric_name = self.choose_basename_for_file(path)
            out_df = self.convert_one_df(df, rule, metric_name)
            tpl = self.edit_export_name.text().strip() or self.rule.get("output_name_template", "清洗_{basename}.xlsx")
            out_name = tpl.format(basename=os.path.splitext(os.path.basename(path))[0])
            out_name = self.ensure_xlsx_ext(out_name)
            out_path = os.path.join(self.export_folder, out_name)
            out_df.to_excel(out_path, index=False, engine="openpyxl")
            self.log(f"成功导出（单文件）：{out_name}")
        except Exception as e:
            self.log(f"导出出错：{e}", error=True)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelCleanerGeneral()
    window.show()
    sys.exit(app.exec())