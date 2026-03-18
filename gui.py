from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import sys
from data_manager import DataManager, ScoringSystem
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib

matplotlib.use('Qt5Agg')

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.data_manager = DataManager()
        self.current_questions = []
        self.current_question_index = 0
        
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("题库随机抽题评分系统")
        self.setGeometry(100, 100, 1200, 800)
        
        # 创建中心部件和主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 第一部分：文件加载和基本信息
        file_group = QGroupBox("题库加载与设置")
        file_layout = QHBoxLayout()
        
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setPlaceholderText("请选择题库Excel文件")
        file_browse_btn = QPushButton("浏览")
        file_browse_btn.clicked.connect(self.browse_file)
        
        file_layout.addWidget(QLabel("题库文件:"))
        file_layout.addWidget(self.file_path_edit)
        file_layout.addWidget(file_browse_btn)
        file_layout.addWidget(QLabel("抽取数量:"))
        
        self.num_questions_spin = QSpinBox()
        self.num_questions_spin.setRange(1, 1000)
        self.num_questions_spin.setValue(10)
        file_layout.addWidget(self.num_questions_spin)
        
        self.load_btn = QPushButton("加载题库")
        self.load_btn.clicked.connect(self.load_question_bank)
        file_layout.addWidget(self.load_btn)
        
        file_group.setLayout(file_layout)
        
        # 第二部分：学生信息和抽题
        student_group = QGroupBox("学生信息与抽题")
        student_layout = QHBoxLayout()
        
        student_layout.addWidget(QLabel("学生姓名:"))
        self.student_name_edit = QLineEdit()
        self.student_name_edit.setPlaceholderText("输入学生姓名")
        student_layout.addWidget(self.student_name_edit)
        
        self.random_btn = QPushButton("随机抽题")
        self.random_btn.clicked.connect(self.random_select_questions)
        self.random_btn.setEnabled(False)
        student_layout.addWidget(self.random_btn)
        
        self.view_history_btn = QPushButton("查看历史")
        self.view_history_btn.clicked.connect(self.view_history)
        student_layout.addWidget(self.view_history_btn)
        
        student_group.setLayout(student_layout)
        
        # 第三部分：题目显示和评分
        question_group = QGroupBox("题目评分")
        question_layout = QVBoxLayout()
        
        # 题目导航
        nav_layout = QHBoxLayout()
        self.prev_btn = QPushButton("上一题")
        self.prev_btn.clicked.connect(self.prev_question)
        self.next_btn = QPushButton("下一题")
        self.next_btn.clicked.connect(self.next_question)
        
        self.question_label = QLabel("题目: 0/0")
        nav_layout.addWidget(self.prev_btn)
        nav_layout.addWidget(self.question_label)
        nav_layout.addWidget(self.next_btn)
        question_layout.addLayout(nav_layout)
        
        # 题目内容
        self.question_text = QTextEdit()
        self.question_text.setReadOnly(True)
        self.question_text.setMaximumHeight(150)
        question_layout.addWidget(QLabel("题目内容:"))
        question_layout.addWidget(self.question_text)
        
        # 评分控件
        scores_layout = QGridLayout()
        
        # 知识点分数
        self.knowledge_slider = QSlider(Qt.Horizontal)
        self.knowledge_slider.setRange(0, 25)
        self.knowledge_slider.valueChanged.connect(self.update_scores)
        self.knowledge_label = QLabel("0")
        scores_layout.addWidget(QLabel("知识点分数 (0-25):"), 0, 0)
        scores_layout.addWidget(self.knowledge_slider, 0, 1)
        scores_layout.addWidget(self.knowledge_label, 0, 2)
        
        # 流利程度
        self.fluency_slider = QSlider(Qt.Horizontal)
        self.fluency_slider.setRange(0, 25)
        self.fluency_slider.valueChanged.connect(self.update_scores)
        self.fluency_label = QLabel("0")
        scores_layout.addWidget(QLabel("回答流利程度 (0-25):"), 1, 0)
        scores_layout.addWidget(self.fluency_slider, 1, 1)
        scores_layout.addWidget(self.fluency_label, 1, 2)
        
        # 知识扩展
        self.expansion_slider = QSlider(Qt.Horizontal)
        self.expansion_slider.setRange(0, 25)
        self.expansion_slider.valueChanged.connect(self.update_scores)
        self.expansion_label = QLabel("0")
        scores_layout.addWidget(QLabel("对知识点的扩展 (0-25):"), 2, 0)
        scores_layout.addWidget(self.expansion_slider, 2, 1)
        scores_layout.addWidget(self.expansion_label, 2, 2)
        
        # 思路严谨
        self.rigor_slider = QSlider(Qt.Horizontal)
        self.rigor_slider.setRange(0, 25)
        self.rigor_slider.valueChanged.connect(self.update_scores)
        self.rigor_label = QLabel("0")
        scores_layout.addWidget(QLabel("回答思路是否严谨 (0-25):"), 3, 0)
        scores_layout.addWidget(self.rigor_slider, 3, 1)
        scores_layout.addWidget(self.rigor_label, 3, 2)
        
        question_layout.addLayout(scores_layout)
        
        # 总分显示
        total_layout = QHBoxLayout()
        total_layout.addWidget(QLabel("当前题目总分:"))
        self.total_label = QLabel("0")
        self.total_label.setStyleSheet("font-size: 16px; font-weight: bold; color: blue;")
        total_layout.addWidget(self.total_label)
        total_layout.addStretch()
        
        question_layout.addLayout(total_layout)
        
        # 评语
        question_layout.addWidget(QLabel("评语:"))
        self.comment_edit = QTextEdit()
        self.comment_edit.setMaximumHeight(80)
        question_layout.addWidget(self.comment_edit)
        
        question_group.setLayout(question_layout)
        
        # 第四部分：统计信息和保存
        stats_group = QGroupBox("统计与保存")
        stats_layout = QHBoxLayout()
        
        # 平均分显示
        avg_layout = QVBoxLayout()
        self.avg_knowledge_label = QLabel("知识点平均分: 0.00")
        self.avg_fluency_label = QLabel("流利程度平均分: 0.00")
        self.avg_expansion_label = QLabel("知识扩展平均分: 0.00")
        self.avg_rigor_label = QLabel("思路严谨平均分: 0.00")
        self.avg_total_label = QLabel("总分平均分: 0.00")
        
        avg_layout.addWidget(self.avg_knowledge_label)
        avg_layout.addWidget(self.avg_fluency_label)
        avg_layout.addWidget(self.avg_expansion_label)
        avg_layout.addWidget(self.avg_rigor_label)
        avg_layout.addWidget(self.avg_total_label)
        
        stats_layout.addLayout(avg_layout)
        
        # 保存按钮
        self.save_btn = QPushButton("保存考核结果")
        self.save_btn.clicked.connect(self.save_results)
        self.save_btn.setEnabled(False)
        stats_layout.addWidget(self.save_btn)
        
        # 图表按钮
        self.chart_btn = QPushButton("显示评分图表")
        self.chart_btn.clicked.connect(self.show_chart)
        self.chart_btn.setEnabled(False)
        stats_layout.addWidget(self.chart_btn)
        
        stats_group.setLayout(stats_layout)
        
        # 将所有组添加到主布局
        main_layout.addWidget(file_group)
        main_layout.addWidget(student_group)
        main_layout.addWidget(question_group)
        main_layout.addWidget(stats_group)
        
        # 状态栏
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        
        # 初始化界面状态
        self.update_question_display()
        
    def browse_file(self):
        """浏览文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择题库文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.file_path_edit.setText(file_path)
            
    def load_question_bank(self):
        """加载题库"""
        file_path = self.file_path_edit.text()
        if not file_path:
            QMessageBox.warning(self, "警告", "请先选择题库文件")
            return
            
        success, message = self.data_manager.load_question_bank(file_path)
        if success:
            self.random_btn.setEnabled(True)
            self.statusBar.showMessage(message, 5000)
        else:
            QMessageBox.critical(self, "错误", message)
            
    def random_select_questions(self):
        """随机抽题"""
        if not self.student_name_edit.text().strip():
            QMessageBox.warning(self, "警告", "请输入学生姓名")
            return
            
        num_questions = self.num_questions_spin.value()
        self.current_questions = self.data_manager.get_random_questions(num_questions)
        
        if self.current_questions:
            self.current_question_index = 0
            self.save_btn.setEnabled(True)
            self.chart_btn.setEnabled(True)
            self.update_question_display()
            self.update_statistics_display()
            self.statusBar.showMessage(f"已抽取 {num_questions} 道题目", 3000)
        else:
            QMessageBox.warning(self, "警告", "题库为空或抽取失败")
            
    def update_question_display(self):
        """更新题目显示"""
        if not self.current_questions:
            self.question_label.setText("题目: 0/0")
            self.question_text.clear()
            return
            
        question = self.current_questions[self.current_question_index]
        
        # 更新题目信息
        self.question_label.setText(
            f"题目: {self.current_question_index + 1}/{len(self.current_questions)}"
        )
        self.question_text.setText(question['问题'])
        
        # 更新评分控件
        self.knowledge_slider.setValue(question['知识点分数'])
        self.fluency_slider.setValue(question['流利程度'])
        self.expansion_slider.setValue(question['知识扩展'])
        self.rigor_slider.setValue(question['思路严谨'])
        
        # 更新评语
        self.comment_edit.setText(question.get('评语', ''))
        
        # 更新按钮状态
        self.prev_btn.setEnabled(self.current_question_index > 0)
        self.next_btn.setEnabled(self.current_question_index < len(self.current_questions) - 1)
        
    def update_scores(self):
        """更新分数显示"""
        if not self.current_questions:
            return
            
        # 获取当前分数
        knowledge_score = self.knowledge_slider.value()
        fluency_score = self.fluency_slider.value()
        expansion_score = self.expansion_slider.value()
        rigor_score = self.rigor_slider.value()
        
        # 更新标签
        self.knowledge_label.setText(str(knowledge_score))
        self.fluency_label.setText(str(fluency_score))
        self.expansion_label.setText(str(expansion_score))
        self.rigor_label.setText(str(rigor_score))
        
        # 计算总分
        total_score = knowledge_score + fluency_score + expansion_score + rigor_score
        self.total_label.setText(str(total_score))
        
        # 更新当前题目数据
        question = self.current_questions[self.current_question_index]
        question['知识点分数'] = knowledge_score
        question['流利程度'] = fluency_score
        question['知识扩展'] = expansion_score
        question['思路严谨'] = rigor_score
        question['总分'] = total_score
        question['评语'] = self.comment_edit.toPlainText()
        
        # 更新统计显示
        self.update_statistics_display()
        
    def update_statistics_display(self):
        """更新统计信息显示"""
        if not self.current_questions:
            return
            
        stats = ScoringSystem.calculate_statistics(self.current_questions)
        
        self.avg_knowledge_label.setText(f"知识点平均分: {stats['avg_knowledge']:.2f}")
        self.avg_fluency_label.setText(f"流利程度平均分: {stats['avg_fluency']:.2f}")
        self.avg_expansion_label.setText(f"知识扩展平均分: {stats['avg_expansion']:.2f}")
        self.avg_rigor_label.setText(f"思路严谨平均分: {stats['avg_rigor']:.2f}")
        self.avg_total_label.setText(f"总分平均分: {stats['avg_total']:.2f}")
        
    def prev_question(self):
        """上一题"""
        if self.current_question_index > 0:
            self.current_question_index -= 1
            self.update_question_display()
            
    def next_question(self):
        """下一题"""
        if self.current_question_index < len(self.current_questions) - 1:
            self.current_question_index += 1
            self.update_question_display()
            
    def save_results(self):
        """保存考核结果"""
        if not self.student_name_edit.text().strip():
            QMessageBox.warning(self, "警告", "请输入学生姓名")
            return
            
        if not self.current_questions:
            QMessageBox.warning(self, "警告", "没有可保存的考核数据")
            return
            
        student_name = self.student_name_edit.text().strip()
        stats = ScoringSystem.calculate_statistics(self.current_questions)
        
        success, message = self.data_manager.save_results(
            student_name, self.current_questions, stats
        )
        
        if success:
            QMessageBox.information(self, "成功", message)
        else:
            QMessageBox.critical(self, "错误", message)
            
    def view_history(self):
        """查看历史记录"""
        student_name = self.student_name_edit.text().strip()
        if not student_name:
            QMessageBox.warning(self, "警告", "请输入学生姓名")
            return
            
        history_df = self.data_manager.load_student_history(student_name)
        
        if history_df.empty:
            QMessageBox.information(self, "提示", f"没有找到{student_name}的历史记录")
            return
            
        # 创建历史记录对话框
        history_dialog = QDialog(self)
        history_dialog.setWindowTitle(f"{student_name}的历史考核记录")
        history_dialog.setGeometry(200, 200, 800, 600)
        
        layout = QVBoxLayout()
        
        # 创建表格显示历史记录
        table = QTableWidget()
        table.setColumnCount(len(history_df.columns))
        table.setHorizontalHeaderLabels(history_df.columns)
        table.setRowCount(len(history_df))
        
        for i in range(len(history_df)):
            for j in range(len(history_df.columns)):
                item = QTableWidgetItem(str(history_df.iloc[i, j]))
                table.setItem(i, j, item)
                
        layout.addWidget(table)
        
        # 关闭按钮
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(history_dialog.close)
        layout.addWidget(close_btn)
        
        history_dialog.setLayout(layout)
        history_dialog.exec_()
        
    def show_chart(self):
        """显示评分图表"""
        if not self.current_questions:
            return
            
        stats = ScoringSystem.calculate_statistics(self.current_questions)
        
        # 创建图表窗口
        chart_dialog = QDialog(self)
        chart_dialog.setWindowTitle("评分统计图表")
        chart_dialog.setGeometry(300, 300, 600, 500)
        
        layout = QVBoxLayout()
        
        # 创建matplotlib图形
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(8, 10))
        
        # 第一个图：各项平均分
        categories = ['知识点分数', '流利程度', '知识扩展', '思路严谨']
        values = [
            stats['avg_knowledge'],
            stats['avg_fluency'],
            stats['avg_expansion'],
            stats['avg_rigor']
        ]
        
        bars = ax1.bar(categories, values, color=['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4'])
        ax1.set_ylabel('平均分')
        ax1.set_title('各项评分平均分')
        ax1.set_ylim(0, 25)
        
        # 在柱状图上显示数值
        for bar, value in zip(bars, values):
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                    f'{value:.2f}', ha='center', va='bottom')
        
        # 第二个图：雷达图
        ax2 = fig.add_subplot(2, 1, 2, projection='polar')
        
        # 数据准备
        categories = ['知识点\n分数', '流利\n程度', '知识\n扩展', '思路\n严谨', '总分']
        values = values + [stats['avg_total'] / 4]  # 将总分转换为25分制
        
        # 雷达图需要闭合数据
        values = values + [values[0]]
        angles = [n / float(len(categories)) * 2 * 3.14159 for n in range(len(categories))]
        angles += angles[:1]
        
        ax2.plot(angles, values, 'o-', linewidth=2)
        ax2.fill(angles, values, alpha=0.25)
        ax2.set_xticks(angles[:-1])
        ax2.set_xticklabels(categories)
        ax2.set_ylim(0, 25)
        ax2.set_title('评分雷达图')
        
        plt.tight_layout()
        
        # 将图形嵌入到Qt窗口
        canvas = FigureCanvas(fig)
        layout.addWidget(canvas)
        
        chart_dialog.setLayout(layout)
        chart_dialog.exec_()
        
        plt.close(fig)