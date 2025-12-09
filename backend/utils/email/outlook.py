"""
Outlook邮件处理模块
"""

import imaplib
import email
import requests
import time

from .common import (
    decode_mime_words,
    normalize_check_time,
    format_date_for_imap_search,
)
from .logger import logger

class OutlookMailHandler:
    """Outlook邮箱处理类"""

    # Outlook常用文件夹映射
    DEFAULT_FOLDERS = {
        'INBOX': ['inbox', 'Inbox', 'INBOX'],
        'SENT': ['sentitems', 'Sent Items', 'Sent', '已发送'],
        'DRAFTS': ['drafts', 'Drafts', '草稿箱'],
        'TRASH': ['deleteditems', 'Deleted Items', 'Trash', '已删除'],
        'SPAM': ['junkemail', 'Junk E-mail', 'Spam', '垃圾邮件'],
        'ARCHIVE': ['archive', 'Archive', '归档']
    }

    def __init__(self, email_address, access_token):
        """初始化Outlook处理器"""
        self.email_address = email_address
        self.access_token = access_token
        self.mail = None
        self.error = None

    def connect(self):
        """连接到Outlook服务器"""
        try:
            self.mail = imaplib.IMAP4_SSL('outlook.live.com')
            auth_string = OutlookMailHandler.generate_auth_string(self.email_address, self.access_token)
            self.mail.authenticate('XOAUTH2', lambda x: auth_string)
            return True
        except Exception as e:
            self.error = str(e)
            logger.error(f"Outlook连接失败: {e}")
            return False

    def get_folders(self):
        """获取文件夹列表"""
        if not self.mail:
            return []

        try:
            _, folders = self.mail.list()
            folder_list = []

            for folder in folders:
                if isinstance(folder, bytes):
                    folder = folder.decode('utf-8', errors='ignore')

                # 解析文件夹名称
                parts = folder.split('"')
                if len(parts) >= 3:
                    folder_name = parts[-2]
                else:
                    folder_name = folder.split()[-1]

                if folder_name and folder_name not in ['.', '..']:
                    folder_list.append(folder_name)

            # 确保常用文件夹在列表中
            default_folders = ['inbox', 'sentitems', 'drafts', 'deleteditems', 'junkemail']
            for df in default_folders:
                if df not in folder_list:
                    folder_list.append(df)

            return sorted(folder_list)
        except Exception as e:
            logger.error(f"获取Outlook文件夹列表失败: {e}")
            return ['inbox']

    def get_messages(self, folder="inbox", limit=100):
        """获取指定文件夹的邮件"""
        if not self.mail:
            return []

        try:
            self.mail.select(folder)
            _, messages = self.mail.search(None, 'ALL')
            message_numbers = messages[0].split()

            # 限制数量并倒序（最新的在前）
            message_numbers = message_numbers[-limit:] if len(message_numbers) > limit else message_numbers
            message_numbers.reverse()

            mail_list = []
            for num in message_numbers:
                try:
                    _, msg_data = self.mail.fetch(num, '(RFC822)')
                    email_body = msg_data[0][1]
                    msg = email.message_from_bytes(email_body)

                    # 简化的邮件解析
                    subject = decode_mime_words(msg.get('Subject', ''))
                    sender = decode_mime_words(msg.get('From', ''))
                    received_time = email.utils.parsedate_to_datetime(msg.get('Date', ''))

                    # 使用统一的新解析逻辑
                    content = OutlookMailHandler._extract_rich_content(msg)

                    mail_list.append({
                        'subject': subject,
                        'sender': sender,
                        'received_time': received_time,
                        'content': content,
                        'folder': folder
                    })
                except Exception as e:
                    logger.warning(f"解析Outlook邮件失败: {e}")
                    continue

            return mail_list
        except Exception as e:
            logger.error(f"获取Outlook邮件失败: {e}")
            return []

    def close(self):
        """关闭连接"""
        if self.mail:
            try:
                self.mail.logout()
            except:
                pass
            self.mail = None

    @staticmethod
    def get_new_access_token(refresh_token, client_id):
        """刷新获取新的access_token"""
        url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
        data = {
            'client_id': client_id,
            'grant_type': 'refresh_token',
            'refresh_token': refresh_token,
        }
        try:
            response = requests.post(url, data=data)
            result_status = response.json().get('error')
            if result_status is not None:
                logger.error(f"获取访问令牌失败: {result_status}")
                return None
            else:
                new_access_token = response.json()['access_token']
                logger.info("成功获取新的访问令牌")
                return new_access_token
        except Exception as e:
            logger.error(f"刷新令牌过程中发生异常: {str(e)}")
            return None

    @staticmethod
    def generate_auth_string(user, token):
        """生成 OAuth2 授权字符串"""
        return f"user={user}\1auth=Bearer {token}\1\1"

    @staticmethod
    def _extract_rich_content(msg):
        """
        辅助方法：解析更丰富的邮件内容（优先HTML，保留附件名）
        """
        text_content = ""
        html_content = ""
        attachments = []
        
        # 1. 遍历邮件结构
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                
                # 获取文件名（如果有）
                filename = part.get_filename()
                if filename:
                    filename = decode_mime_words(filename)
                    attachments.append(filename)
                
                # 如果是附件，跳过内容解析
                if 'attachment' in content_disposition:
                    continue

                # 解析正文内容
                try:
                    payload = part.get_payload(decode=True)
                    if not payload: continue
                    
                    # 尝试探测字符集，默认utf-8
                    charset = part.get_content_charset()
                    if not charset:
                        charset = 'utf-8'
                    
                    decoded_part = payload.decode(charset, errors='replace')

                    if content_type == 'text/html':
                        html_content += decoded_part
                    elif content_type == 'text/plain':
                        text_content += decoded_part
                except Exception:
                    pass
        else:
            # 非多部分邮件（通常是纯文本）
            try:
                payload = msg.get_payload(decode=True)
                charset = msg.get_content_charset() or 'utf-8'
                text_content = payload.decode(charset, errors='replace')
            except:
                text_content = str(msg.get_payload())

        # 2. 组装最终内容
        # 策略：如果有HTML，优先使用HTML（保留格式），否则使用纯文本
        final_content = html_content if html_content.strip() else text_content
        
        # 3. 将附件信息追加到正文底部
        if attachments:
            att_str = "<br><hr><b>[系统提示] 包含附件:</b> " + ", ".join(attachments)
            if not html_content.strip():
                # 如果是纯文本，用文本方式追加
                att_str = f"\n\n----------------\n[包含附件]: {', '.join(attachments)}"
            final_content += att_str

        return final_content

    @staticmethod
    def fetch_emails(email_address, access_token, folder="inbox", callback=None, last_check_time=None):
        """
        通过IMAP协议获取Outlook/Hotmail邮件
        逻辑变更：
        1. 自动查找垃圾邮件并全部移动到收件箱
        2. 只从收件箱获取邮件
        """
        mail_records = []

        if callback is None:
            callback = lambda progress, folder: None

        last_check_time = normalize_check_time(last_check_time)
        
        # 常见垃圾箱名称
        junk_aliases = ['Junk Email', 'Junk', 'Spam', '垃圾邮件', 'junkemail']

        logger.info(f"开始处理账户 {email_address}")

        max_retries = 3
        
        for retry in range(max_retries):
            try:
                callback(10, "连接服务器...")
                # 建立连接
                mail = imaplib.IMAP4_SSL('outlook.live.com')
                auth_string = OutlookMailHandler.generate_auth_string(email_address, access_token)
                mail.authenticate('XOAUTH2', lambda x: auth_string)
                
                # --- 第一步：将垃圾邮件移动到收件箱 ---
                callback(20, "检查垃圾邮件...")
                for junk_name in junk_aliases:
                    try:
                        status, _ = mail.select(junk_name)
                        if status == 'OK':
                            logger.info(f"发现垃圾邮件文件夹: {junk_name}，准备迁移...")
                            
                            # 获取所有垃圾邮件ID
                            status, data = mail.search(None, 'ALL')
                            if status == 'OK':
                                mail_ids = data[0].split()
                                if mail_ids:
                                    logger.info(f"正在移动 {len(mail_ids)} 封垃圾邮件到收件箱")
                                    # IMAP copy 需要逗号分隔的ID序列
                                    id_set = b','.join(mail_ids)
                                    
                                    # 1. 复制到收件箱
                                    res, _ = mail.copy(id_set, 'INBOX')
                                    if res == 'OK':
                                        # 2. 标记原邮件为删除
                                        mail.store(id_set, '+FLAGS', '\\Deleted')
                                        # 3. 永久删除 (Expunge)
                                        mail.expunge()
                                        logger.info("垃圾邮件迁移完成")
                                    else:
                                        logger.error(f"复制垃圾邮件失败: {res}")
                            # 找到一个有效的垃圾箱处理完后就停止查找别名
                            break 
                    except Exception as e:
                        # 忽略单个文件夹错误，继续尝试下一个别名
                        continue

                # --- 第二步：从收件箱获取所有邮件 ---
                callback(40, "正在获取收件箱...")
                mail.select('INBOX')

                if last_check_time:
                    search_date = format_date_for_imap_search(last_check_time)
                    search_cmd = f'(SINCE "{search_date}")'
                    logger.info(f"搜索 {search_date} 之后的邮件")
                    status, data = mail.search(None, search_cmd)
                else:
                    # 获取最近的100封
                    status, data = mail.search(None, 'ALL')

                if status != 'OK':
                    logger.error("无法搜索收件箱")
                    mail.logout()
                    return []

                mail_ids = data[0].split()
                # 限制处理最新的 100 封
                mail_ids = mail_ids[-100:] if len(mail_ids) > 100 else mail_ids
                
                total_mails = len(mail_ids)
                logger.info(f"收件箱中待处理邮件: {total_mails}")

                for i, mail_id in enumerate(mail_ids):
                    # 进度 40% - 90%
                    progress = 40 + int((i / total_mails) * 50) if total_mails else 90
                    callback(progress, "INBOX")

                    try:
                        _, mail_data = mail.fetch(mail_id, '(RFC822)')
                        msg = email.message_from_bytes(mail_data[0][1])

                        subject = decode_mime_words(msg.get('Subject', ''))
                        sender = decode_mime_words(msg.get('From', ''))
                        received_time = email.utils.parsedate_to_datetime(msg.get('Date', ''))

                        # 生成唯一标识
                        mail_key = f"{subject}|{sender}|{received_time.isoformat() if received_time else 'unknown'}"

                        # 内存去重
                        if mail_key in [record.get('mail_key') for record in mail_records]:
                            continue

                        # 使用富文本解析
                        content = OutlookMailHandler._extract_rich_content(msg)

                        mail_records.append({
                            'subject': subject,
                            'sender': sender,
                            'received_time': received_time,
                            'content': content,
                            'mail_key': mail_key,
                            'folder': 'INBOX' 
                        })

                    except Exception as e:
                        logger.error(f"解析邮件ID {mail_id} 失败: {e}")

                callback(90, "完成获取")
                # 成功后退出重试循环
                break

            except Exception as e:
                logger.error(f"IMAP操作异常 (尝试 {retry+1}/{max_retries}): {str(e)}")
                time.sleep(1)
            finally:
                try:
                    mail.logout()
                except:
                    pass

        return mail_records

    @staticmethod
    def check_mail(email_info, db, progress_callback=None):
        """检查Outlook/Hotmail邮箱中的邮件并存储到数据库"""
        email_id = email_info['id']
        email_address = email_info['email']
        refresh_token = email_info['refresh_token']
        client_id = email_info['client_id']

        logger.info(f"开始检查Outlook邮箱: ID={email_id}, 邮箱={email_address}")

        # 确保回调函数存在
        if progress_callback is None:
            progress_callback = lambda progress, message: None

        # 报告初始进度
        progress_callback(0, "正在获取访问令牌...")

        try:
            # 获取新的访问令牌
            access_token = OutlookMailHandler.get_new_access_token(refresh_token, client_id)
            if not access_token:
                error_msg = f"邮箱{email_address}(ID={email_id})获取访问令牌失败"
                logger.error(error_msg)
                progress_callback(0, error_msg)
                return {
                    'success': False,
                    'message': error_msg
                }

            # 更新令牌到数据库
            db.update_email_token(email_id, access_token)

            # 报告进度
            progress_callback(10, "开始获取邮件...")

            # 获取邮件
            def folder_progress_callback(progress, folder):
                msg = f"正在处理{folder}，进度{progress}%"
                # 将内部进度映射到总进度10-90%
                total_progress = 10 + int(progress * 0.8)
                progress_callback(total_progress, msg)

            try:
                # 调用 fetch_emails (内部已经包含了移动垃圾邮件和抓取Inbox的逻辑)
                mail_records = OutlookMailHandler.fetch_emails(
                    email_address,
                    access_token,
                    "inbox", 
                    folder_progress_callback
                )

                # 报告进度
                count = len(mail_records)
                progress_callback(90, f"获取到{count}封邮件，正在保存...")

                # 将邮件记录保存到数据库
                saved_count = 0
                for record in mail_records:
                    try:
                        success = db.add_mail_record(
                            email_id,
                            record['subject'],
                            record['sender'],
                            record['received_time'],
                            record['content']
                        )
                        if success:
                            saved_count += 1
                    except Exception as e:
                        logger.error(f"保存邮件记录失败: {str(e)}")

                # 更新最后检查时间
                try:
                    # 只要流程没报错就更新时间
                    db.update_check_time(email_id)
                    logger.info(f"已更新邮箱{email_address}(ID={email_id})的最后检查时间")
                except Exception as e:
                    logger.error(f"更新检查时间失败: {str(e)}")

                # 报告完成
                success_msg = f"完成，共处理{count}封邮件，新增{saved_count}封"
                progress_callback(100, success_msg)

                logger.info(f"邮箱{email_address}(ID={email_id})检查完成，获取到{count}封邮件，新增{saved_count}封")
                return {
                    'success': True,
                    'message': success_msg,
                    'total': count,
                    'saved': saved_count
                }

            except Exception as e:
                error_msg = f"检查邮件失败: {str(e)}"
                logger.error(f"邮箱{email_address}(ID={email_id}){error_msg}")
                progress_callback(0, error_msg)
                return {
                    'success': False,
                    'message': error_msg
                }

        except Exception as e:
            error_msg = f"处理邮箱过程中出错: {str(e)}"
            logger.error(f"邮箱{email_address}(ID={email_id}){error_msg}")
            progress_callback(0, error_msg)
            return {
                'success': False,
                'message': error_msg
            }
