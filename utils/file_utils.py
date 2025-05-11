import os
import json
import re
from datetime import datetime, timedelta

def save_credentials(email, password):
    """保存邮箱和密码到配置文件"""
    config = {
        'email': email,
        'password': password,
        'failed_attempts': 0,
        'lock_time': None
    }
    with open('config.json', 'w') as f:
        json.dump(config, f)

def load_credentials():
    """从配置文件加载邮箱和密码"""
    try:
        with open('config.json') as f:
            return json.load(f)
    except:
        return None

def validate_email(email):
    """验证邮箱格式"""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

def validate_password(password):
    """验证密码长度"""
    return len(password) >= 6

def is_locked(config):
    """检查是否处于锁定状态"""
    if config.get('lock_time'):
        lock_time = datetime.strptime(config['lock_time'], '%Y-%m-%d %H:%M:%S')
        return datetime.now() < lock_time + timedelta(minutes=15)
    return False
import shutil
import hashlib
import smtplib
from email.mime.text import MIMEText
from datetime import datetime

class FileUtils:
    def __init__(self):
        self.password_hash = None
        self.user_email = None
        self.attempt_count = 0
        self.locked_until = None
        
    def set_credentials(self, password, email):
        """Set password and email for security"""
        self.password_hash = hashlib.sha256(password.encode()).hexdigest()
        self.user_email = email
        
    def verify_password(self, password):
        """Verify password with lockout after 3 attempts"""
        if self.locked_until and datetime.now() < self.locked_until:
            return False
            
        if hashlib.sha256(password.encode()).hexdigest() == self.password_hash:
            self.attempt_count = 0
            return True
            
        self.attempt_count += 1
        if self.attempt_count >= 3:
            self.lock_system()
            self.send_alert_email("Too many password attempts - system locked")
        return False
        
    def lock_system(self):
        """Lock system for 15 minutes"""
        self.locked_until = datetime.now() + timedelta(minutes=15)
        
    def send_alert_email(self, message):
        """Send security alert email"""
        if not self.user_email:
            return False
            
        try:
            msg = MIMEText(message)
            msg['Subject'] = 'Product Logger Security Alert'
            msg['From'] = 'noreply@productlogger.com'
            msg['To'] = self.user_email
            
            # TODO: Configure SMTP settings
            with smtplib.SMTP('localhost') as server:
                server.send_message(msg)
            return True
        except Exception:
            return False
            
    def upload_file(self, file_path, target_dir):
        """Handle file upload with backup and organization"""
        try:
            # Create backup
            backup_dir = os.path.join(target_dir, 'backup')
            os.makedirs(backup_dir, exist_ok=True)
            backup_path = os.path.join(backup_dir, os.path.basename(file_path))
            shutil.copy2(file_path, backup_path)
            
            # Organize by file type
            ext = os.path.splitext(file_path)[1].lower()[1:]  # Remove dot
            type_dir = os.path.join(target_dir, ext)
            os.makedirs(type_dir, exist_ok=True)
            
            # Save with timestamp and original filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            original_name = os.path.splitext(os.path.basename(file_path))[0]
            new_filename = f"{timestamp}_{original_name}.{ext}"
            final_path = os.path.join(type_dir, new_filename)
            
            # Remove old files if exists (only one allowed per type)
            for f in os.listdir(type_dir):
                if f.endswith(f".{ext}"):
                    os.remove(os.path.join(type_dir, f))
                
            shutil.move(file_path, final_path)
            return final_path
        except Exception as e:
            return None
            
    def get_file_info(self, file_path):
        """Get file information for display"""
        try:
            size = os.path.getsize(file_path)
            stats = {
                'filename': os.path.basename(file_path),
                'size': f"{size/1024:.2f} KB",
                'modified': datetime.fromtimestamp(os.path.getmtime(file_path))
            }
            
            if file_path.lower().endswith(('.xlsx', '.csv')):
                df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)
                stats['rows'] = len(df)
                stats['columns'] = list(df.columns)
                
            return stats
        except Exception:
            return None
