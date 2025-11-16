import sys
import random
import string
import imaplib
import email
import smtplib
import os
from email.message import EmailMessage
import requests
from email.header import decode_header

from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException

# 全局配置
OUTLOOK_IMAP_SERVER = "imap-mail.outlook.com"
OUTLOOK_IMAP_PORT = 993
CODE_KEYWORDS = ['验证', '验证码', '注册码', 'Verification', 'Verification Code', 'Registration Code']

class TempEmailService:
    """临时邮箱注册服务"""
    
    @staticmethod
    def generate_email_prefix():
        """生成邮箱前缀（首字母+7位字母数字）"""
        first_char = random.choice(string.ascii_lowercase)
        rest = ''.join(random.choices(string.ascii_lowercase + string.digits, k=7))
        return first_char + rest

    @staticmethod
    def register_outlook():
        """注册Outlook邮箱（启动 Edge 自动填表，返回账号信息，人工完成验证码）"""
        driver = None
        try:
            # 1. 生成邮箱和密码
            prefix = TempEmailService.generate_email_prefix()
            domain = random.choice(["outlook.com", "hotmail.com"])
            email_addr = f"{prefix}@{domain}"
            password = ''.join(random.choices(
                string.ascii_letters + string.digits + '!@#$%', k=12))

            # 2. 配置 Edge 浏览器
            options = webdriver.EdgeOptions()
            # 如需无头运行可打开下一行
            # options.add_argument("--headless=new")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.use_chromium = True

            driver_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "msedgedriver.exe")
            if not os.path.exists(driver_path):
                raise FileNotFoundError(f"未找到浏览器驱动: {driver_path}")

            driver = webdriver.Edge(
                service=Service(driver_path),
                options=options
            )
            driver.set_page_load_timeout(30)

            # 3. 打开注册页面
            driver.get("https://signup.live.com/signup")

            # 可能出现“同意并继续”弹窗
            try:
                agree_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//button[contains(., '同意并继续')]")
                    )
                )
                agree_btn.click()
            except TimeoutException:
                # 没有这个弹窗就继续
                pass

            # 4. 输入邮箱地址
            try:
                email_input = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.NAME, "电子邮件"))
                )
            except TimeoutException:
                # 备用：根据 id 定位
                email_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "floatingLabelInput7"))
                )

            email_input.send_keys(email_addr)

            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(., '下一步')]")
                )
            ).click()

            # 5. 输入密码
            try:
                pwd_input = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.ID, "floatingLabelInput16"))
                )
            except TimeoutException:
                pwd_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.NAME, "Password"))
                )

            pwd_input.send_keys(password)

            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(., '下一步')]")
                )
            ).click()

            # 6. 填写国家/地区和生日（如果页面出现）
            try:
                country_select = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.ID, "Country"))
                )
                country_select.send_keys("中国")

                year_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.NAME, "BirthYear"))
                )
                year_input.clear()
                year_input.send_keys("1995")

                month_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "BirthMonthDropdown"))
                )
                month_btn.click()
                month_option = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//*[@role='option' and contains(., '1月')]")
                    )
                )
                month_option.click()

                day_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "BirthDayDropdown"))
                )
                day_btn.click()
                day_option = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//*[@role='option' and contains(., '1日')]")
                    )
                )
                day_option.click()

                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//button[contains(., '下一步')]")
                    )
                ).click()
            except TimeoutException:
                # 有时界面略有变化，尽量不中断流程
                pass

            # 7. 填写姓名（如果页面出现）
            try:
                first_name = ''.join(random.choices(string.ascii_lowercase, k=5)).capitalize()
                last_name = ''.join(random.choices(string.ascii_lowercase, k=6)).capitalize()

                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.NAME, "FirstName"))
                ).send_keys(first_name)
                driver.find_element(By.NAME, "LastName").send_keys(last_name)
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//button[contains(., '下一步')]")
                    )
                ).click()
            except TimeoutException:
                pass

            # 接下来通常进入"证明你不是机器人"页面，此处交给用户手动完成
            return {
                "success": True,
                "email": email_addr,
                "password": password,
                "message": "已自动完成前置步骤，请在浏览器中完成人机验证，然后在界面中点击验证按钮。"
            }
        except (TimeoutException, WebDriverException) as e:
            return {"success": False, "message": f"浏览器自动化出错：{e}"}
        except Exception as e:
            return {"success": False, "message": str(e)}
        finally:
            # 不主动关闭浏览器，方便用户继续完成人机验证
            pass

    @staticmethod
    def register_mail_tm():
        """注册mail.tm临时邮箱"""
        try:
            prefix = TempEmailService.generate_email_prefix()
            # 获取域名
            domains_resp = requests.get("https://api.mail.tm/domains", timeout=10)
            if domains_resp.status_code != 200:
                return {"success": False, "message": f"获取域名失败：{domains_resp.status_code}"}

            domains_data = domains_resp.json()
            members = domains_data.get("hydra:member", [])
            if not members:
                return {"success": False, "message": "无可用域名"}

            domain = members[0].get("domain")
            email_addr = f"{prefix}@{domain}"
            password = ''.join(random.choices(
                string.ascii_letters + string.digits + '!@#$%', k=12))

            # 创建账号
            acc_resp = requests.post(
                "https://api.mail.tm/accounts",
                json={"address": email_addr, "password": password},
                timeout=10,
            )
            if acc_resp.status_code not in (200, 201):
                return {"success": False, "message": f"创建失败：{acc_resp.text}"}

            # 获取token
            token_resp = requests.post(
                "https://api.mail.tm/token",
                json={"address": email_addr, "password": password},
                timeout=10,
            )
            token_data = token_resp.json()
            token = token_data.get("token")
            
            return {
                "success": True,
                "email": email_addr,
                "password": password,
                "token": token
            }
        except Exception as e:
            return {"success": False, "message": str(e)}

    @staticmethod
    def register_1secmail():
        """注册1secmail临时邮箱"""
        try:
            resp = requests.get(
                "https://www.1secmail.com/api/v1/",
                params={"action": "genRandomMailbox", "count": 1},
                timeout=10,
            )
            if resp.status_code != 200:
                return {"success": False, "message": f"创建失败：{resp.status_code}"}

            data = resp.json()
            if not data:
                return {"success": False, "message": "未返回邮箱地址"}

            return {
                "success": True,
                "email": data[0],
                "password": ""
            }
        except Exception as e:
            return {"success": False, "message": str(e)}

    @staticmethod
    def register_guerrillamail():
        """注册GuerrillaMail临时邮箱"""
        try:
            resp = requests.get(
                "https://api.guerrillamail.com/ajax.php",
                params={"f": "get_email_address"},
                timeout=10,
            )
            if resp.status_code != 200:
                return {"success": False, "message": f"创建失败：{resp.status_code}"}

            data = resp.json()
            email_addr = data.get("email_addr")
            sid_token = data.get("sid_token")
            
            if not email_addr or not sid_token:
                return {"success": False, "message": "信息不完整"}

            return {
                "success": True,
                "email": email_addr,
                "sid_token": sid_token,
                "password": ""
            }
        except Exception as e:
            return {"success": False, "message": str(e)}


class EmailHandler:
    """邮箱收发处理服务"""
    
    @staticmethod
    def fetch_verification_code(email_info):
        """获取验证码相关邮件内容"""
        email_type = email_info.get("type")
        
        try:
            if email_type == "outlook":
                return EmailHandler._fetch_outlook_code(email_info)
            elif email_type == "mail.tm":
                return EmailHandler._fetch_mail_tm_code(email_info)
            elif email_type == "1secmail":
                return EmailHandler._fetch_1secmail_code(email_info)
            elif email_type == "guerrillamail":
                return EmailHandler._fetch_guerrillamail_code(email_info)
            else:
                return {"success": False, "message": "未知邮箱类型"}
        except Exception as e:
            return {"success": False, "message": str(e)}

    @staticmethod
    def _fetch_outlook_code(email_info):
        """获取Outlook邮箱验证码"""
        mail = imaplib.IMAP4_SSL(OUTLOOK_IMAP_SERVER, OUTLOOK_IMAP_PORT)
        mail.login(email_info["email"], email_info["password"])
        mail.select('inbox')

        # 搜索关键词
        search_criteria = []
        for keyword in CODE_KEYWORDS:
            search_criteria.extend(['OR', 'SUBJECT', keyword, 'BODY', keyword])
        search_criteria = search_criteria[1:] if search_criteria else []

        status, data = mail.search(None, *search_criteria)
        if status != 'OK' or not data[0]:
            return {"success": False, "message": "未找到相关邮件"}

        latest_email_id = data[0].split()[-1]
        status, msg_data = mail.fetch(latest_email_id, '(RFC822)')
        msg = email.message_from_bytes(msg_data[0][1])

        # 提取内容
        email_content = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    email_content = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                    break
        else:
            email_content = msg.get_payload(decode=True).decode('utf-8', errors='ignore')

        mail.logout()
        return EmailHandler._extract_code_content(email_content)

    @staticmethod
    def _fetch_mail_tm_code(email_info):
        """获取mail.tm邮箱验证码"""
        headers = {"Authorization": f"Bearer {email_info.get('token')}"}
        msg_list_resp = requests.get("https://api.mail.tm/messages", headers=headers, timeout=10)
        
        if msg_list_resp.status_code != 200:
            return {"success": False, "message": f"获取邮件列表失败：{msg_list_resp.status_code}"}

        msg_list = msg_list_resp.json().get("hydra:member", [])
        if not msg_list:
            return {"success": False, "message": "未找到邮件"}

        latest_msg = msg_list[-1]
        msg_resp = requests.get(
            f"https://api.mail.tm/messages/{latest_msg.get('id')}",
            headers=headers,
            timeout=10
        )
        
        msg_data = msg_resp.json()
        email_content = msg_data.get("text") or "\n".join(msg_data.get("html", [])) or ""
        return EmailHandler._extract_code_content(email_content)

    @staticmethod
    def _fetch_1secmail_code(email_info):
        """获取1secmail邮箱验证码"""
        try:
            login, domain = email_info["email"].split('@', 1)
        except ValueError:
            return {"success": False, "message": "邮箱格式错误"}

        # 获取邮件列表
        params = {"action": "getMessages", "login": login, "domain": domain}
        resp = requests.get("https://www.1secmail.com/api/v1/", params=params, timeout=10)
        if resp.status_code != 200:
            return {"success": False, "message": f"获取列表失败：{resp.status_code}"}

        msg_list = resp.json()
        if not msg_list:
            return {"success": False, "message": "未找到邮件"}

        # 获取最新邮件
        latest = msg_list[-1]
        detail_params = {
            "action": "readMessage",
            "login": login,
            "domain": domain,
            "id": latest.get("id"),
        }
        detail_resp = requests.get("https://www.1secmail.com/api/v1/", params=detail_params, timeout=10)
        
        msg_data = detail_resp.json()
        email_content = msg_data.get("textBody") or msg_data.get("body") or ""
        return EmailHandler._extract_code_content(email_content)

    @staticmethod
    def _fetch_guerrillamail_code(email_info):
        """获取GuerrillaMail邮箱验证码"""
        params = {
            "f": "check_email", 
            "sid_token": email_info.get("sid_token"), 
            "seq": 0
        }
        resp = requests.get("https://api.guerrillamail.com/ajax.php", params=params, timeout=10)
        if resp.status_code != 200:
            return {"success": False, "message": f"获取列表失败：{resp.status_code}"}

        data = resp.json()
        msg_list = data.get("list", [])
        if not msg_list:
            return {"success": False, "message": "未找到邮件"}

        # 获取最新邮件
        latest = msg_list[0]
        detail_params = {
            "f": "fetch_email", 
            "sid_token": email_info.get("sid_token"), 
            "email_id": latest.get("mail_id")
        }
        detail_resp = requests.get("https://api.guerrillamail.com/ajax.php", params=detail_params, timeout=10)
        
        msg_data = detail_resp.json()
        email_content = msg_data.get("mail_body", "")
        return EmailHandler._extract_code_content(email_content)

    @staticmethod
    def _extract_code_content(email_content):
        """提取验证码相关内容"""
        for keyword in CODE_KEYWORDS:
            if keyword in email_content:
                idx = email_content.index(keyword)
                start = max(0, idx - 100)
                end = min(len(email_content), idx + len(keyword) + 100)
                return {
                    "success": True,
                    "related": email_content[start:end].strip(),
                    "full": email_content
                }
        
        return {
            "success": True,
            "related": "找到关键词邮件，但未提取到相关核心内容",
            "full": email_content
        }


class MailboxService:
    """邮箱授权码收发服务（QQ/163等）"""
    
    @staticmethod
    def get_server_info(mail_type):
        """获取邮箱服务器信息"""
        servers = {
            "qq": {
                "imap": "imap.qq.com",
                "imap_port": 993,
                "smtp": "smtp.qq.com",
                "smtp_port": 465
            },
            "163": {
                "imap": "imap.163.com",
                "imap_port": 993,
                "smtp": "smtp.163.com",
                "smtp_port": 465
            }
        }
        return servers.get(mail_type, {})

    @staticmethod
    def test_connection(mail_type, email, password):
        """测试邮箱连接"""
        server_info = MailboxService.get_server_info(mail_type)
        if not server_info:
            return False, "未知邮箱类型"

        try:
            with imaplib.IMAP4_SSL(server_info["imap"], server_info["imap_port"]) as mail:
                mail.login(email, password)
                mail.select('inbox')
            return True, "连接成功"
        except Exception as e:
            return False, str(e)

    @staticmethod
    def fetch_mail_list(mail_type, email, password, page=0, page_size=10):
        """获取邮件列表"""
        server_info = MailboxService.get_server_info(mail_type)
        if not server_info:
            return False, "未知邮箱类型", []

        try:
            with imaplib.IMAP4_SSL(server_info["imap"], server_info["imap_port"]) as mail:
                mail.login(email, password)
                mail.select('inbox')
                
                # 获取所有邮件ID
                status, data = mail.search(None, 'ALL')
                if status != 'OK':
                    return False, "获取邮件列表失败", []

                all_ids = data[0].split()
                total = len(all_ids)
                start = max(0, total - (page + 1) * page_size)
                end = max(0, total - page * page_size)
                page_ids = all_ids[start:end][::-1]  # 倒序显示

                # 获取邮件摘要
                mail_list = []
                for msg_id in page_ids:
                    status, msg_data = mail.fetch(msg_id, '(RFC822.HEADER)')
                    msg = email.message_from_bytes(msg_data[0][1])
                    
                    # 解析主题
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding or "utf-8")

                    # 解析发件人
                    from_addr = msg.get("From", "")
                    
                    mail_list.append({
                        "id": msg_id.decode(),
                        "subject": subject,
                        "from": from_addr,
                        "date": msg.get("Date", "")
                    })

                return True, f"共 {total} 封邮件", mail_list, total
        except Exception as e:
            return False, str(e), [], 0

    @staticmethod
    def get_mail_content(mail_type, email, password, msg_id):
        """获取邮件内容"""
        server_info = MailboxService.get_server_info(mail_type)
        if not server_info:
            return False, "未知邮箱类型", ""

        try:
            with imaplib.IMAP4_SSL(server_info["imap"], server_info["imap_port"]) as mail:
                mail.login(email, password)
                mail.select('inbox')
                
                status, msg_data = mail.fetch(msg_id, '(RFC822)')
                msg = email.message_from_bytes(msg_data[0][1])

                content = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        if content_type in ("text/plain", "text/html"):
                            payload = part.get_payload(decode=True)
                            encoding = part.get_content_charset() or "utf-8"
                            content += payload.decode(encoding, errors="ignore") + "\n"
                else:
                    payload = msg.get_payload(decode=True)
                    encoding = msg.get_content_charset() or "utf-8"
                    content = payload.decode(encoding, errors="ignore")

                return True, "获取成功", content
        except Exception as e:
            return False, str(e), ""

    @staticmethod
    def send_email(mail_type, email, password, to_addr, subject, content):
        """发送邮件"""
        server_info = MailboxService.get_server_info(mail_type)
        if not server_info:
            return False, "未知邮箱类型"

        try:
            msg = EmailMessage()
            msg["From"] = email
            msg["To"] = to_addr
            msg["Subject"] = subject
            msg.set_content(content)

            with smtplib.SMTP_SSL(server_info["smtp"], server_info["smtp_port"]) as server:
                server.login(email, password)
                server.send_message(msg)
            
            return True, "发送成功"
        except Exception as e:
            return False, str(e)