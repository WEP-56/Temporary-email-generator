import sys
import json
import os
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QPushButton, QListWidget, QTextEdit, QTextBrowser, QLabel, QMessageBox, 
                               QLineEdit, QComboBox, QTabWidget)
from PySide6.QtCore import Qt, QThread, Signal
from email_services import TempEmailService, EmailHandler, MailboxService

# 全局配置
EMAIL_LIST_PATH = "email_list.json"
MAIL_ACCOUNTS_PATH = "mail_accounts.json"

class EmailRegisterThread(QThread):
    """邮箱注册线程"""
    finish_signal = Signal(dict)

    def __init__(self, email_type):
        super().__init__()
        self.email_type = email_type

    def run(self):
        try:
            if self.email_type == "outlook":
                result = TempEmailService.register_outlook()
            elif self.email_type == "mail.tm":
                result = TempEmailService.register_mail_tm()
            elif self.email_type == "1secmail":
                result = TempEmailService.register_1secmail()
            elif self.email_type == "guerrillamail":
                result = TempEmailService.register_guerrillamail()
            else:
                result = {"success": False, "message": "未知邮箱类型"}
            
            result["type"] = self.email_type
            self.finish_signal.emit(result)
        except Exception as e:
            self.finish_signal.emit({
                "success": False,
                "message": str(e),
                "type": self.email_type
            })


class EmailCheckThread(QThread):
    """邮箱验证码查询线程"""
    code_signal = Signal(dict)

    def __init__(self, email_info):
        super().__init__()
        self.email_info = email_info

    def run(self):
        try:
            result = EmailHandler.fetch_verification_code(self.email_info)
            self.code_signal.emit(result)
        except Exception as e:
            self.code_signal.emit({"success": False, "message": str(e)})


class EmailRegisterApp(QMainWindow):
    """主界面类"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("邮箱工具")
        self.setGeometry(100, 100, 800, 600)
        self.email_list = self.load_email_list()
        self.mail_accounts = self.load_mail_accounts()
        self.pending_outlook_account = None
        self.init_ui()

    def init_ui(self):
        """初始化界面"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        outer_layout = QVBoxLayout(central_widget)

        self.tabs = QTabWidget()
        outer_layout.addWidget(self.tabs)

        # 邮箱注册器标签页
        self.init_register_tab()
        # 邮件收发器标签页
        self.init_mailbox_tab()

    def init_register_tab(self):
        """初始化注册标签页"""
        tab1 = QWidget()
        tab1_layout = QHBoxLayout(tab1)

        # 左侧布局
        left_layout = QVBoxLayout()
        self.email_list_widget = QListWidget()
        self.refresh_email_list()
        self.email_list_widget.itemClicked.connect(self.on_email_clicked)
        
        left_layout.addWidget(QLabel("已注册邮箱列表"))
        left_layout.addWidget(self.email_list_widget)

        # 操作按钮
        btn_row = QHBoxLayout()
        self.mark_used_btn = QPushButton("标记已使用")
        self.mark_used_btn.clicked.connect(self.mark_email_used)
        self.delete_btn = QPushButton("删除邮箱")
        self.delete_btn.clicked.connect(self.delete_email)
        btn_row.addWidget(self.mark_used_btn)
        btn_row.addWidget(self.delete_btn)
        left_layout.addLayout(btn_row)

        # 查询按钮
        self.query_btn = QPushButton("查询验证码")
        self.query_btn.clicked.connect(self.query_selected_email)
        left_layout.addWidget(self.query_btn)

        # 邮箱类型选择
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("邮箱类型："))
        self.email_type_combo = QComboBox()
        self.email_type_combo.addItem("Outlook邮箱", "outlook")
        self.email_type_combo.addItem("临时邮箱 mail.tm", "mail.tm")
        self.email_type_combo.addItem("临时邮箱 1secmail", "1secmail")
        self.email_type_combo.addItem("临时邮箱 GuerrillaMail", "guerrillamail")
        type_layout.addWidget(self.email_type_combo)
        left_layout.addLayout(type_layout)

        # 注册按钮
        self.register_btn = QPushButton("新建邮箱")
        self.register_btn.clicked.connect(self.start_register)
        left_layout.addWidget(self.register_btn)

        # Outlook验证按钮
        self.verify_btn = QPushButton("我已完成验证（Outlook）")
        self.verify_btn.setEnabled(False)
        self.verify_btn.clicked.connect(self.verify_outlook_registration)
        left_layout.addWidget(self.verify_btn)

        # 右侧布局
        right_layout = QVBoxLayout()
        self.related_label = QLabel("关键词相关内容（可直接复制）：")
        self.related_display = QTextBrowser()
        self.related_display.setOpenExternalLinks(True)
        
        self.copy_btn = QPushButton("复制相关内容")
        self.copy_btn.clicked.connect(self.copy_related)
        
        self.log_display = QTextEdit()
        self.log_display.setReadOnly(True)
        self.log_display.setPlaceholderText("操作日志...")

        right_layout.addWidget(self.related_label)
        right_layout.addWidget(self.related_display)
        right_layout.addWidget(self.copy_btn)
        right_layout.addWidget(QLabel("日志："))
        right_layout.addWidget(self.log_display)

        tab1_layout.addLayout(left_layout, stretch=1)
        tab1_layout.addLayout(right_layout, stretch=2)
        self.tabs.addTab(tab1, "邮箱注册器")

    def init_mailbox_tab(self):
        """初始化邮件收发标签页"""
        tab2 = QWidget()
        tab2_layout = QVBoxLayout(tab2)

        # 账号配置区
        acct_layout = QHBoxLayout()
        acct_layout.addWidget(QLabel("邮箱类型："))
        self.mail_type_combo = QComboBox()
        self.mail_type_combo.addItem("QQ 邮箱", "qq")
        self.mail_type_combo.addItem("163 邮箱", "163")
        self.mail_type_combo.currentIndexChanged.connect(self.apply_account_to_fields)
        acct_layout.addWidget(self.mail_type_combo)

        acct_layout.addWidget(QLabel("邮箱地址："))
        self.mail_email_edit = QLineEdit()
        acct_layout.addWidget(self.mail_email_edit)

        acct_layout.addWidget(QLabel("密码 / 授权码："))
        self.mail_pass_edit = QLineEdit()
        self.mail_pass_edit.setEchoMode(QLineEdit.Password)
        acct_layout.addWidget(self.mail_pass_edit)

        self.mail_test_btn = QPushButton("测试连接")
        self.mail_test_btn.clicked.connect(self.test_mail_account)
        acct_layout.addWidget(self.mail_test_btn)

        tab2_layout.addLayout(acct_layout)
        tab2_layout.addWidget(QLabel("提示：QQ/163 建议在邮箱设置中开启 IMAP/SMTP，并使用授权码登录。"))

        # 中部：收件箱列表和正文
        mid_layout = QHBoxLayout()

        left_mail_layout = QVBoxLayout()
        self.mail_fetch_btn = QPushButton("刷新收件箱")
        self.mail_fetch_btn.clicked.connect(self.fetch_mailbox)
        left_mail_layout.addWidget(self.mail_fetch_btn)
        self.mail_list_widget = QListWidget()
        self.mail_list_widget.itemClicked.connect(self.on_mail_item_clicked)
        left_mail_layout.addWidget(self.mail_list_widget)

        # 分页按钮
        mail_page_btn_row = QHBoxLayout()
        self.mail_first_page_btn = QPushButton("回到首页")
        self.mail_prev_page_btn = QPushButton("上一页")
        self.mail_next_page_btn = QPushButton("下一页")
        self.mail_first_page_btn.clicked.connect(self.goto_mail_first_page)
        self.mail_prev_page_btn.clicked.connect(self.goto_mail_prev_page)
        self.mail_next_page_btn.clicked.connect(self.goto_mail_next_page)
        mail_page_btn_row.addWidget(self.mail_first_page_btn)
        mail_page_btn_row.addWidget(self.mail_prev_page_btn)
        mail_page_btn_row.addWidget(self.mail_next_page_btn)
        left_mail_layout.addLayout(mail_page_btn_row)

        right_mail_layout = QVBoxLayout()
        right_mail_layout.addWidget(QLabel("邮件内容："))
        self.mail_body_display = QTextBrowser()
        self.mail_body_display.setOpenExternalLinks(True)
        right_mail_layout.addWidget(self.mail_body_display)

        mid_layout.addLayout(left_mail_layout, stretch=1)
        mid_layout.addLayout(right_mail_layout, stretch=2)
        tab2_layout.addLayout(mid_layout)

        # 发信区域
        send_layout = QVBoxLayout()
        row_to = QHBoxLayout()
        row_to.addWidget(QLabel("收件人："))
        self.mail_send_to_edit = QLineEdit()
        row_to.addWidget(self.mail_send_to_edit)
        send_layout.addLayout(row_to)

        row_subject = QHBoxLayout()
        row_subject.addWidget(QLabel("主题："))
        self.mail_send_subject_edit = QLineEdit()
        row_subject.addWidget(self.mail_send_subject_edit)
        send_layout.addLayout(row_subject)

        send_layout.addWidget(QLabel("正文："))
        self.mail_send_body_edit = QTextEdit()
        send_layout.addWidget(self.mail_send_body_edit)

        self.mail_send_btn = QPushButton("发送邮件")
        self.mail_send_btn.clicked.connect(self.send_mail)
        send_layout.addWidget(self.mail_send_btn)

        tab2_layout.addLayout(send_layout)
        self.tabs.addTab(tab2, "邮件收发器")

        # 初始化分页状态
        self._mail_page_size = 10
        self._mail_current_page = 0
        self._all_mail_count = 0
        self._current_mail_ids = []
        self.apply_account_to_fields()

    # 数据加载与保存
    def load_email_list(self):
        if os.path.exists(EMAIL_LIST_PATH):
            with open(EMAIL_LIST_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        return []

    def load_mail_accounts(self):
        if os.path.exists(MAIL_ACCOUNTS_PATH):
            try:
                with open(MAIL_ACCOUNTS_PATH, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def save_mail_accounts(self):
        try:
            with open(MAIL_ACCOUNTS_PATH, 'w', encoding='utf-8') as f:
                json.dump(self.mail_accounts, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.warning(self, "提示", f"保存邮箱账号信息失败：{e}")

    def save_email_list(self):
        with open(EMAIL_LIST_PATH, 'w', encoding='utf-8') as f:
            json.dump(self.email_list, f, ensure_ascii=False, indent=2)

    # 邮箱列表操作
    def refresh_email_list(self):
        self.email_list_widget.clear()
        for idx, item in enumerate(self.email_list):
            status = "已使用" if item.get("used") else "未使用"
            self.email_list_widget.addItem(
                f"[{idx + 1}] {item['email']} (密码：{item.get('password','')}) ({status})"
            )

    def on_email_clicked(self, item):
        """选中邮箱项"""
        pass

    def mark_email_used(self):
        """标记邮箱为已使用"""
        current_idx = self.email_list_widget.currentRow()
        if current_idx >= 0 and current_idx < len(self.email_list):
            self.email_list[current_idx]["used"] = True
            self.save_email_list()
            self.refresh_email_list()
            self.append_log(f"标记邮箱为已使用：{self.email_list[current_idx]['email']}")

    def delete_email(self):
        """删除邮箱"""
        current_idx = self.email_list_widget.currentRow()
        if current_idx >= 0 and current_idx < len(self.email_list):
            email = self.email_list.pop(current_idx)
            self.save_email_list()
            self.refresh_email_list()
            self.append_log(f"删除邮箱：{email['email']}")

    # 邮箱注册相关
    def start_register(self):
        """开始注册邮箱"""
        email_type = self.email_type_combo.currentData()
        self.register_btn.setEnabled(False)
        self.verify_btn.setEnabled(False)
        
        self.register_thread = EmailRegisterThread(email_type)
        self.register_thread.finish_signal.connect(self.on_register_finish)
        self.register_thread.start()
        self.append_log(f"开始注册{email_type}邮箱...")

    def on_register_finish(self, result):
        """注册完成回调"""
        self.register_btn.setEnabled(True)
        if result["success"]:
            self.append_log(f"{result['type']}邮箱注册成功：{result.get('email')}")
            if result["type"] == "outlook":
                self.pending_outlook_account = {
                    "email": result["email"],
                    "password": result["password"],
                    "type": "outlook",
                    "used": False
                }
                self.verify_btn.setEnabled(True)
                QMessageBox.information(
                    self,
                    "提示",
                    f"请在浏览器完成人机验证后点击验证按钮\n邮箱：{result['email']}\n密码：{result['password']}"
                )
            else:
                # 构建邮箱信息字典
                email_info = {
                    "email": result["email"],
                    "password": result.get("password", ""),
                    "type": result["type"],
                    "used": False
                }
                # 添加特定类型的额外信息
                if result["type"] == "mail.tm":
                    email_info["token"] = result.get("token")
                elif result["type"] == "guerrillamail":
                    email_info["sid_token"] = result.get("sid_token")

                self.email_list.append(email_info)
                self.save_email_list()
                self.refresh_email_list()
                QMessageBox.information(
                    self,
                    "成功",
                    f"{result['type']}邮箱注册成功\n邮箱：{result['email']}"
                )
        else:
            self.append_log(f"{result['type']}邮箱注册失败：{result['message']}")
            QMessageBox.warning(self, "失败", result["message"])

    def verify_outlook_registration(self):
        """验证Outlook注册"""
        if self.pending_outlook_account:
            self.email_list.append(self.pending_outlook_account)
            self.save_email_list()
            self.refresh_email_list()
            self.append_log(f"Outlook邮箱验证完成：{self.pending_outlook_account['email']}")
            self.verify_btn.setEnabled(False)
            self.pending_outlook_account = None
            QMessageBox.information(self, "成功", "邮箱已添加到列表")

    # 验证码查询
    def query_selected_email(self):
        """查询选中邮箱的验证码"""
        current_idx = self.email_list_widget.currentRow()
        if current_idx < 0 or current_idx >= len(self.email_list):
            QMessageBox.warning(self, "提示", "请先选择一个邮箱")
            return

        email_info = self.email_list[current_idx]
        self.append_log(f"开始查询{email_info['email']}的验证码...")
        
        self.check_thread = EmailCheckThread(email_info)
        self.check_thread.code_signal.connect(self.on_code_received)
        self.check_thread.start()

    def on_code_received(self, result):
        """收到验证码回调"""
        if result["success"]:
            self.related_display.setText(result["related"])
            self.append_log("验证码查询成功")
        else:
            self.append_log(f"验证码查询失败：{result['message']}")
            QMessageBox.warning(self, "失败", result["message"])

    # 邮件收发相关
    def apply_account_to_fields(self):
        """应用保存的账号到输入框"""
        mail_type = self.mail_type_combo.currentData()
        if mail_type in self.mail_accounts:
            self.mail_email_edit.setText(self.mail_accounts[mail_type].get("email", ""))
            self.mail_pass_edit.setText(self.mail_accounts[mail_type].get("password", ""))

    def test_mail_account(self):
        """测试邮箱连接"""
        mail_type = self.mail_type_combo.currentData()
        email = self.mail_email_edit.text()
        password = self.mail_pass_edit.text()

        if not email or not password:
            QMessageBox.warning(self, "提示", "请填写邮箱和密码")
            return

        # 保存账号信息
        self.mail_accounts[mail_type] = {"email": email, "password": password}
        self.save_mail_accounts()

        # 测试连接
        success, msg = MailboxService.test_connection(mail_type, email, password)
        if success:
            QMessageBox.information(self, "成功", msg)
            self.append_log(f"邮箱连接测试成功：{email}")
        else:
            QMessageBox.warning(self, "失败", msg)
            self.append_log(f"邮箱连接测试失败：{msg}")

    def fetch_mailbox(self):
        """刷新收件箱"""
        mail_type = self.mail_type_combo.currentData()
        email = self.mail_email_edit.text()
        password = self.mail_pass_edit.text()

        if not email or not password:
            QMessageBox.warning(self, "提示", "请填写邮箱和密码")
            return

        self.append_log("正在获取邮件列表...")
        success, msg, mail_list, total = MailboxService.fetch_mail_list(
            mail_type, email, password, self._mail_current_page, self._mail_page_size
        )

        if success:
            self._all_mail_count = total
            self.mail_list_widget.clear()
            self._current_mail_ids = []
            for mail in mail_list:
                self._current_mail_ids.append(mail["id"])
                self.mail_list_widget.addItem(f"{mail['subject']} - {mail['from']}")
            self.append_log(msg)
        else:
            QMessageBox.warning(self, "失败", msg)
            self.append_log(f"获取邮件列表失败：{msg}")

    def on_mail_item_clicked(self, item):
        """点击邮件项查看内容"""
        idx = self.mail_list_widget.row(item)
        if idx < 0 or idx >= len(self._current_mail_ids):
            return

        mail_id = self._current_mail_ids[idx]
        mail_type = self.mail_type_combo.currentData()
        email = self.mail_email_edit.text()
        password = self.mail_pass_edit.text()

        success, msg, content = MailboxService.get_mail_content(
            mail_type, email, password, mail_id
        )

        if success:
            self.mail_body_display.setText(content)
        else:
            QMessageBox.warning(self, "失败", msg)
            self.append_log(f"获取邮件内容失败：{msg}")

    def send_mail(self):
        """发送邮件"""
        mail_type = self.mail_type_combo.currentData()
        email = self.mail_email_edit.text()
        password = self.mail_pass_edit.text()
        to_addr = self.mail_send_to_edit.text()
        subject = self.mail_send_subject_edit.text()
        content = self.mail_send_body_edit.toPlainText()

        if not to_addr or not subject or not content:
            QMessageBox.warning(self, "提示", "请填写收件人、主题和正文")
            return

        success, msg = MailboxService.send_email(
            mail_type, email, password, to_addr, subject, content
        )

        if success:
            QMessageBox.information(self, "成功", msg)
            self.append_log(f"邮件发送成功：{to_addr}")
            # 清空发送框
            self.mail_send_to_edit.clear()
            self.mail_send_subject_edit.clear()
            self.mail_send_body_edit.clear()
        else:
            QMessageBox.warning(self, "失败", msg)
            self.append_log(f"邮件发送失败：{msg}")

    # 分页控制
    def goto_mail_first_page(self):
        """回到首页"""
        if self._mail_current_page != 0:
            self._mail_current_page = 0
            self.fetch_mailbox()

    def goto_mail_prev_page(self):
        """上一页"""
        if self._mail_current_page > 0:
            self._mail_current_page -= 1
            self.fetch_mailbox()

    def goto_mail_next_page(self):
        """下一页"""
        max_page = (self._all_mail_count - 1) // self._mail_page_size
        if self._mail_current_page < max_page:
            self._mail_current_page += 1
            self.fetch_mailbox()

    # 辅助功能
    def append_log(self, content):
        """添加日志"""
        self.log_display.append(content)
        # 滚动到底部
        self.log_display.verticalScrollBar().setValue(
            self.log_display.verticalScrollBar().maximum()
        )

    def copy_related(self):
        """复制相关内容"""
        content = self.related_display.toPlainText()
        if content:
            clipboard = QApplication.clipboard()
            clipboard.setText(content)
            QMessageBox.information(self, "提示", "内容已复制到剪贴板")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = EmailRegisterApp()
    window.show()
    sys.exit(app.exec())