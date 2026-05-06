#!/usr/bin/env python3
"""
step3_send.py - 简报邮件发送脚本

支持发送 .md 和 .docx 格式的简报文件。
- .md 文件：自动转换为 HTML 正文发送
- .docx 文件：作为附件发送，正文自动生成摘要
- 混合使用：--brief 简报.md --attach 简报.docx 同时发送正文和附件

用法：
  python3 step3_send.py --brief 简报文件路径
  python3 step3_send.py --brief 简报文件路径 --attach 附件路径
  python3 step3_send.py --brief data/日期/brief/xxx.docx
  python3 step3_send.py --dir data/2026-04-25/brief/
  python3 step3_send.py --dir data/2026-04-25/brief/ --config email.json
"""

import json
import os
import sys
import smtplib
import argparse
import mimetypes
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email.header import Header
from email import encoders
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_DIR = os.path.join(BASE_DIR, "config")

DEFAULT_EMAIL_CONFIG = os.path.join(CONFIG_DIR, "email.json")


def load_email_config(config_path=None):
    """加载邮件配置"""
    if config_path is None:
        config_path = DEFAULT_EMAIL_CONFIG
    if not os.path.exists(config_path):
        print(f"❌ 邮件配置文件不存在: {config_path}")
        print("请先创建 config/email.json，格式：")
        print("""{
  "smtp_server": "smtp.163.com",
  "smtp_port": 465,
  "use_ssl": true,
  "sender": "youraccount@163.com",
  "password": "SMTP授权码",
  "recipients": ["a@qq.com", "b@163.com"],
  "subject_prefix": "[行业日报]"
}""")
        sys.exit(1)
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def read_file(filepath, binary=False):
    """读取文件内容"""
    if not os.path.exists(filepath):
        print(f"❌ 文件不存在: {filepath}")
        return None
    mode = 'rb' if binary else 'r'
    encoding = None if binary else 'utf-8'
    with open(filepath, mode, encoding=encoding) as f:
        return f.read()


def send_email(email_cfg, subject, body_html, body_text=None, attachments=None):
    """
    发送邮件（支持附件）
    
    attachments: list of dict, 每个元素含 {"path": str, "filename": str (可选)}
    """
    smtp_server = email_cfg["smtp_server"]
    smtp_port = email_cfg.get("smtp_port", 465)
    use_ssl = email_cfg.get("use_ssl", True)
    sender = email_cfg["sender"]
    password = email_cfg["password"]
    recipients = email_cfg["recipients"]
    subject_prefix = email_cfg.get("subject_prefix", "[行业日报]")

    full_subject = f"{subject_prefix}{subject}"

    # 有附件时用 mixed，否则用 alternative
    if attachments:
        msg = MIMEMultipart('mixed')
    else:
        msg = MIMEMultipart('alternative')

    msg['From'] = sender
    msg['To'] = ', '.join(recipients)
    msg['Subject'] = Header(full_subject, 'utf-8')

    # 正文部分
    body_related = MIMEMultipart('related')
    body_alternative = MIMEMultipart('alternative')

    # 纯文本版本
    if body_text:
        text_part = MIMEText(body_text, 'plain', 'utf-8')
        body_alternative.attach(text_part)

    # HTML 版本
    if body_html:
        html_part = MIMEText(body_html, 'html', 'utf-8')
        body_alternative.attach(html_part)

    body_related.attach(body_alternative)
    msg.attach(body_related)

    # 附件
    if attachments:
        for att in attachments:
            att_path = att.get("path", "")
            att_filename = att.get("filename", os.path.basename(att_path))

            if not os.path.exists(att_path):
                print(f"  ⚠️  附件不存在: {att_path}")
                continue

            try:
                with open(att_path, 'rb') as f:
                    attachment = MIMEApplication(f.read())

                attachment.add_header(
                    'Content-Disposition', 'attachment',
                    filename=('utf-8', '', att_filename)
                )
                msg.attach(attachment)
                print(f"  📎 附件: {att_filename}")
            except Exception as e:
                print(f"  ⚠️  附件添加失败: {e}")

    # 发送
    try:
        if use_ssl:
            server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30)
        else:
            server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
            server.starttls()

        server.login(sender, password)
        server.sendmail(sender, recipients, msg.as_string())
        server.quit()

        print(f"✅ 发送成功: {full_subject}")
        print(f"   收件人: {', '.join(recipients)}")
        return True
    except Exception as e:
        print(f"❌ 发送失败: {e}")
        return False


def md_to_html(md_text):
    """简单的 md 转 html（不依赖第三方库）"""
    import re
    lines = md_text.split('\n')
    html_parts = []
    in_list = False

    for line in lines:
        # 标题
        if line.startswith('# '):
            if in_list:
                html_parts.append('</ul>')
                in_list = False
            html_parts.append(f'<h1>{line[2:]}</h1>')
        elif line.startswith('## '):
            if in_list:
                html_parts.append('</ul>')
                in_list = False
            html_parts.append(f'<h2>{line[3:]}</h2>')
        elif line.startswith('### '):
            if in_list:
                html_parts.append('</ul>')
                in_list = False
            html_parts.append(f'<h3>{line[4:]}</h3>')
        # 无序列表
        elif line.strip().startswith('- ') or line.strip().startswith('* '):
            if not in_list:
                html_parts.append('<ul>')
                in_list = True
            content = line.strip()[2:]
            html_parts.append(f'<li>{content}</li>')
        # 有序列表
        elif re.match(r'^\d+\.\s', line.strip()):
            if not in_list:
                html_parts.append('<ol>')
                in_list = True
            content = re.sub(r'^\d+\.\s', '', line.strip())
            html_parts.append(f'<li>{content}</li>')
        # 空行
        elif not line.strip():
            if in_list:
                html_parts.append('</ul>' if not in_list else '</ol>')
                in_list = False
            html_parts.append('<br>')
        else:
            if in_list:
                html_parts.append('</ul>')
                in_list = False
            # 链接处理
            line = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'<a href="\2">\1</a>', line)
            html_parts.append(f'<p>{line}</p>')

    if in_list:
        html_parts.append('</ul>')

    return '<div style="font-family: Arial, sans-serif; line-height: 1.6;">' + \
           '\n'.join(html_parts) + '</div>'


def send_brief_file(filepath, email_cfg, attach_path=None):
    """发送单个简报文件（支持 .md 和 .docx）"""
    ext = os.path.splitext(filepath)[1].lower()
    filename = os.path.basename(filepath)
    subject = filename.replace('.md', '').replace('.docx', '').replace('_', ' ')

    attachments = []
    body_text = None
    body_html = None

    if ext == '.md':
        # .md 文件：作为正文发送
        brief_text = read_file(filepath)
        if not brief_text:
            return False
        body_text = brief_text
        body_html = md_to_html(brief_text)

    elif ext == '.docx':
        # .docx 文件：作为附件发送，正文生成简短说明
        if not os.path.exists(filepath):
            print(f"❌ 文件不存在: {filepath}")
            return False

        # 将 .docx 作为附件
        attachments.append({"path": filepath, "filename": filename})

        # 尝试从 .docx 中提取文本作为正文预览
        try:
            from docx import Document
            doc = Document(filepath)
            preview_lines = []
            char_count = 0
            for p in doc.paragraphs:
                text = p.text.strip()
                if text:
                    preview_lines.append(text)
                    char_count += len(text)
                    if char_count > 500:
                        preview_lines.append("...(详情见附件Word文档)")
                        break
            body_text = '\n'.join(preview_lines) if preview_lines else f"请查看附件 {filename}"
            body_html = f"<pre style='font-family:微软雅黑, sans-serif; line-height:1.6;'>{body_text}</pre>"
            print(f"  📄 从 .docx 提取 {char_count}+ 字符作为正文预览")
        except ImportError:
            body_text = f"请查看附件 Word 文档: {filename}"
            body_html = f"<p>请查看附件 Word 文档: {filename}</p>"
        except Exception as e:
            print(f"  ⚠️  .docx 预览提取失败: {e}")
            body_text = f"请查看附件 Word 文档: {filename}"
            body_html = f"<p>请查看附件 Word 文档: {filename}</p>"

    else:
        print(f"❌ 不支持的文件格式: {ext}，仅支持 .md / .docx")
        return False

    # 如果有额外的附件参数
    if attach_path:
        if os.path.exists(attach_path):
            attachments.append({"path": attach_path, "filename": os.path.basename(attach_path)})
        else:
            print(f"  ⚠️  指定附件不存在: {attach_path}")

    return send_email(email_cfg, subject, body_html, body_text, attachments if attachments else None)


def send_brief_dir(dirpath, email_cfg):
    """发送目录下所有简报（.md 和 .docx 都支持）"""
    if not os.path.isdir(dirpath):
        print(f"❌ 目录不存在: {dirpath}")
        return False

    # 优先发送 .docx，有 .docx 时对应的 .md 作为附件一起发
    docx_files = sorted([f for f in os.listdir(dirpath) if f.endswith('.docx')])
    md_files = sorted([f for f in os.listdir(dirpath) if f.endswith('.md')])

    all_files = docx_files + [f for f in md_files if f.replace('.md', '.docx') not in docx_files]

    if not all_files:
        print(f"⚠️  目录中没有简报文件: {dirpath}")
        return False

    success = True
    for f in all_files:
        fpath = os.path.join(dirpath, f)
        print(f"\n发送: {f}")

        # 如果是 .docx，看同名的 .md 是否存在并作为额外附件
        attach_path = None
        if f.endswith('.docx'):
            md_peer = f.replace('.docx', '.md')
            md_peer_path = os.path.join(dirpath, md_peer)
            if os.path.exists(md_peer_path):
                attach_path = md_peer_path

        if not send_brief_file(fpath, email_cfg, attach_path=attach_path):
            success = False

    return success


def log_send(date_str, files_sent, success):
    """记录发送日志"""
    sent_dir = os.path.join(BASE_DIR, "data", date_str, "sent")
    os.makedirs(sent_dir, exist_ok=True)
    log_path = os.path.join(sent_dir, "send_log.json")

    log_entry = {
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "files": files_sent,
        "success": success
    }

    if os.path.exists(log_path):
        with open(log_path, "r", encoding="utf-8") as f:
            logs = json.load(f)
    else:
        logs = []

    logs.append(log_entry)
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(logs, f, ensure_ascii=False, indent=2)

    print(f"📝 发送日志已保存: {log_path}")


def main():
    parser = argparse.ArgumentParser(description="简报邮件发送")
    parser.add_argument("--brief", help="简报文件路径")
    parser.add_argument("--dir", help="简报目录路径")
    parser.add_argument("--config", help="邮件配置文件路径（默认 config/email.json）")
    parser.add_argument("--date", help="日期 YYYY-MM-DD（日志用，默认当天）")
    parser.add_argument("--attach", help="额外附件路径")
    args = parser.parse_args()

    email_cfg = load_email_config(args.config)
    date_str = args.date or datetime.now().strftime("%Y-%m-%d")

    if args.brief:
        success = send_brief_file(args.brief, email_cfg, attach_path=args.attach)
        log_send(date_str, [os.path.basename(args.brief)], success)
    elif args.dir:
        success = send_brief_dir(args.dir, email_cfg)
        files_log = sorted(os.listdir(args.dir))
        log_send(date_str, files_log, success)
    else:
        print("请指定 --brief 或 --dir")
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
