#!/usr/bin/env python3
"""
OneDrive 同步工具 - GUI 版本 (优化版)
基于 Microsoft Graph API + CustomTkinter

优化内容说明：
1. Windows 11 风格界面 - 深蓝配色、圆角卡片
2. 顶部蓝色标题栏 - 带云图标
3. 账号管理优化 - 点击选择账号，不再输入数字
4. 状态指示灯 - 彩色圆点实时反馈
5. 同步按钮优化 - 大色块主按钮，上传下载分开
6. 日志显示优化 - 带时间戳和条数统计
7. 添加本地文件夹显示
8. 非Windows系统兼容
"""

import os
import sys
import json
import time
import hashlib
import base64
import shutil
import secrets
import logging
from logging.handlers import RotatingFileHandler
import platform
import subprocess
import webbrowser
import requests
import threading
import queue
import pathlib
import datetime
import typing
import urllib.parse
import http.server
import socketserver

# 单实例检查 (Windows) - 优化: 添加异常处理，兼容非Windows
import ctypes
import ctypes.wintypes

ERROR_ALREADY_EXISTS = 183
MUTEX_NAME = "OneDriveSync_SingleInstance_Mutex"
try:
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if ctypes.windll.kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
        ctypes.windll.kernel32.CloseHandle(mutex)
        print("程序已在运行中，请勿重复启动")
        sys.exit(0)
except:
    pass  # 非Windows系统忽略

# GUI 库
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

# 设置主题 - 优化: 改用 dark-blue 主题
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

# ============== 配色方案 ==============
class Theme:
    """Windows 风格配色"""
    BG_MAIN = "#202020"
    BG_CARD = "#2D2D2D"
    ACCENT = "#0078D4"
    ACCENT_HOVER = "#1E8FFF"
    SUCCESS = "#107C10"
    WARNING = "#FF8C00"  # 橙色
    ERROR = "#E81123"
    TEXT_WHITE = "#FFFFFF"
    TEXT_GRAY = "#9E9E9E"
    TEXT_MUTED = "#6E6E6E"
    BORDER = "#3D3D3D"


# ============== 配置路径 ==============
if platform.system() == "Windows":
    DEFAULT_APP_DATA = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "OneDriveSync")
else:
    DEFAULT_APP_DATA = os.path.expanduser("~/.onedrive_sync")

# 启动位置文件只负责记住“配置目录在哪里”，不包含账号密钥。
BOOTSTRAP_FILE = os.path.join(DEFAULT_APP_DATA, "config_location.json")

def _load_config_dir() -> str:
    os.makedirs(DEFAULT_APP_DATA, exist_ok=True)
    try:
        if os.path.exists(BOOTSTRAP_FILE):
            with open(BOOTSTRAP_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            path = os.path.abspath(os.path.expanduser(data.get('config_dir', '')))
            if path:
                os.makedirs(path, exist_ok=True)
                return path
    except Exception as e:
        logger.warning(f"配置位置读取失败，将使用默认位置: {e}") if 'logger' in globals() else None
    return DEFAULT_APP_DATA

def set_config_dir(path: str):
    """保存新的配置目录位置，并刷新运行时路径。"""
    global APP_DATA, CONFIG_FILE, STATE_FILE
    path = os.path.abspath(os.path.expanduser(path))
    os.makedirs(path, exist_ok=True)
    os.makedirs(DEFAULT_APP_DATA, exist_ok=True)
    tmp = BOOTSTRAP_FILE + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump({'config_dir': path}, f, indent=2, ensure_ascii=False)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, BOOTSTRAP_FILE)
    APP_DATA = path
    CONFIG_FILE = os.path.join(APP_DATA, "config.json")
    STATE_FILE = os.path.join(APP_DATA, "sync_state.json")
    _configure_logging()

APP_DATA = _load_config_dir()
CONFIG_FILE = os.path.join(APP_DATA, "config.json")
STATE_FILE = os.path.join(APP_DATA, "sync_state.json")
os.makedirs(APP_DATA, exist_ok=True)

# ============== 日志 ==============
def _configure_logging():
    """根据当前 APP_DATA 重新绑定日志文件。切换配置目录后不会继续写旧目录。"""
    logger_obj = logging.getLogger("OneDriveSync")
    logger_obj.setLevel(logging.INFO)
    logger_obj.propagate = False
    for handler in list(logger_obj.handlers):
        try:
            handler.flush()
            handler.close()
        finally:
            logger_obj.removeHandler(handler)
    os.makedirs(APP_DATA, exist_ok=True)
    handler = RotatingFileHandler(os.path.join(APP_DATA, 'sync.log'), maxBytes=5 * 1024 * 1024, backupCount=3, encoding='utf-8')
    handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger_obj.addHandler(handler)
    return logger_obj

logger = _configure_logging()


# ============== API 客户端类 (保持不变) ==============

REGION_ENDPOINTS = {
    "global": {"auth": "https://login.microsoftonline.com", "graph": "https://graph.microsoft.com"},
    "cn": {"auth": "https://login.partner.microsoftonline.cn", "graph": "https://microsoftgraph.chinacloudapi.cn"},
    "de": {"auth": "https://login.microsoftonline.de", "graph": "https://graph.microsoft.de"},
    "us": {"auth": "https://login.microsoftonline.us", "graph": "https://graph.microsoft.us"}
}

ACCOUNT_TEMPLATE = {
    "name": "",
    "region": "global",
    "client_id": "",
    "client_secret": "",
    "tenant_id": "common",
    "redirect_uri": "http://localhost:53682/callback",
    "refresh_token": "",
    "user_email": "",
    "root_folder": "/",
    "local_folder": ""
}


class GraphAPIError(RuntimeError):
    """结构化 Graph API 错误，保留 HTTP 状态码和 Retry-After。"""
    def __init__(self, status_code: int, message: str, *, response=None, retry_after: float = 0):
        super().__init__(message)
        self.status_code = status_code
        self.retry_after = retry_after
        self.response = response


class OneDriveClient:
    """OneDrive API 客户端：统一重试、错误处理和路径编码。"""

    RETRY_STATUS = {429, 500, 502, 503, 504}
    MAX_RETRIES = 4

    def __init__(self, config: typing.Dict):
        self.config = config
        self.region = config.get("region", "global")
        self.endpoints = REGION_ENDPOINTS.get(self.region, REGION_ENDPOINTS["global"])
        self.access_token = None
        self.token_expires_at = 0
        self.last_auth_error = ""
        self.last_graph_error = ""

        # 本程序是桌面/本地 GUI 应用，Azure 应用注册应配置为 Public Client。
        # Public Client 不能在 token endpoint 提交 client_secret。
        # 保留配置中的 client_secret 字段只是为了兼容旧配置，不会发送给 Microsoft。

    @property
    def auth_base(self):
        return self.endpoints["auth"]

    @property
    def graph_base(self):
        return self.endpoints["graph"] + "/v1.0"

    @staticmethod
    def _encode_path(path: str) -> str:
        """按路径段编码，避免 #、%、?、空格、Unicode 等文件名破坏 Graph URL。"""
        return "/".join(urllib.parse.quote(part, safe="") for part in path.strip('/').split('/') if part != '')

    @staticmethod
    def _retry_delay(response, attempt: int) -> float:
        retry_after = response.headers.get('Retry-After') if response is not None else None
        if retry_after:
            try:
                return min(max(float(retry_after), 0.5), 60.0)
            except ValueError:
                pass
        return min(2 ** attempt, 16) + (0.1 * attempt)

    def load_token(self) -> bool:
        token = self.config.get("refresh_token")
        if not token:
            return False
        return self.refresh_token()

    def get_authorization_url(self) -> str:
        state = secrets.token_urlsafe(32)
        params = {
            'client_id': self.config['client_id'],
            'response_type': 'code',
            'redirect_uri': self.config['redirect_uri'],
            'scope': 'openid profile email User.Read Files.ReadWrite.All offline_access',
            'response_mode': 'query',
            'state': state
        }
        url = f"{self.auth_base}/{self.config['tenant_id']}/oauth2/v2.0/authorize?{urllib.parse.urlencode(params)}"
        return url, state

    def exchange_code(self, code: str) -> bool:
        """公共客户端使用 authorization code 换取 access/refresh token。

        本程序属于桌面 Public Client，因此 token 请求绝不能携带
        client_secret/client_assertion。否则 Microsoft Entra 会返回 AADSTS700025。
        """
        self.last_auth_error = ""
        code = (code or "").strip()
        if not code:
            self.last_auth_error = "授权码为空"
            return False

        tenant_id = (self.config.get("tenant_id") or "common").strip()
        client_id = (self.config.get("client_id") or "").strip()
        redirect_uri = (self.config.get("redirect_uri") or "").strip()

        if not client_id:
            self.last_auth_error = "Client ID 为空"
            return False
        if not redirect_uri:
            self.last_auth_error = "Redirect URI 为空"
            return False

        url = f"{self.auth_base}/{tenant_id}/oauth2/v2.0/token"
        data = {
            "client_id": client_id,
            "code": code,
            "redirect_uri": redirect_uri,
            "grant_type": "authorization_code",
        }

        try:
            response = requests.post(
                url,
                data=data,
                headers={"Accept": "application/json"},
                timeout=30,
            )
            try:
                payload = response.json()
            except ValueError:
                payload = {}

            if response.status_code >= 400:
                error = payload.get("error", "HTTP_ERROR")
                description = str(payload.get("error_description", "")).strip()
                error_codes = payload.get("error_codes")
                suffix = f" | error_codes={error_codes}" if error_codes else ""
                self.last_auth_error = f"{error}: {description or response.text[:500]}{suffix}"
                logger.error(
                    "授权码交换失败: HTTP %s, error=%s, description=%s, error_codes=%s, redirect_uri=%s",
                    response.status_code,
                    error,
                    description,
                    error_codes,
                    redirect_uri,
                )
                return False

            access_token = payload.get("access_token", "")
            if not access_token:
                self.last_auth_error = "Token 响应中没有 access_token"
                logger.error("Token 响应缺少 access_token")
                return False

            refresh_token = payload.get("refresh_token", "")
            if not refresh_token:
                self.last_auth_error = "Token 响应中没有 refresh_token；请确认授权请求包含 offline_access"
                logger.error("Token 响应缺少 refresh_token")
                return False

            self.config["refresh_token"] = refresh_token
            self.access_token = access_token
            self.token_expires_at = time.time() + int(payload.get("expires_in", 3600) or 3600)
            logger.info("公共客户端授权码交换 Token 成功")
            return True

        except requests.RequestException as e:
            self.last_auth_error = f"网络请求失败: {e}"
            logger.exception("授权码交换网络异常")
            return False
        except Exception as e:
            self.last_auth_error = f"授权处理异常: {e}"
            logger.exception("授权码交换异常")
            return False

    def refresh_token(self) -> bool:
        """公共客户端刷新 Token；同样不能提交 client_secret。"""
        refresh = (self.config.get('refresh_token') or '').strip()
        if not refresh:
            return False

        tenant_id = (self.config.get('tenant_id') or 'common').strip()
        client_id = (self.config.get('client_id') or '').strip()
        if not client_id:
            logger.error("Token 刷新失败: Client ID 为空")
            return False

        url = f"{self.auth_base}/{tenant_id}/oauth2/v2.0/token"
        data = {
            'client_id': client_id,
            'refresh_token': refresh,
            'grant_type': 'refresh_token',
            'scope': 'Files.ReadWrite.All offline_access'
        }
        try:
            response = requests.post(
                url,
                data=data,
                headers={"Accept": "application/json"},
                timeout=30,
            )
            try:
                payload = response.json()
            except ValueError:
                payload = {}

            if response.status_code >= 400:
                error = payload.get('error', 'HTTP_ERROR')
                description = str(payload.get('error_description', '')).strip()
                logger.error(
                    "Token 刷新失败: HTTP %s, error=%s, description=%s, error_codes=%s",
                    response.status_code,
                    error,
                    description,
                    payload.get('error_codes'),
                )
                if response.status_code in (400, 401):
                    logger.warning("Refresh Token 已失效，需要重新授权")
                    self.config['refresh_token'] = ''
                    self.access_token = None
                return False

            access_token = payload.get('access_token')
            if not access_token:
                logger.error("Token 刷新响应缺少 access_token")
                return False

            self.access_token = access_token
            self.token_expires_at = time.time() + int(payload.get('expires_in', 3600) or 3600)
            if payload.get('refresh_token'):
                self.config['refresh_token'] = payload['refresh_token']
            return True
        except (requests.RequestException, ValueError, TypeError) as e:
            logger.error(f"Token 刷新失败: {e}")
            return False

    def _request(self, method: str, endpoint: str, **kwargs) -> typing.Optional[typing.Dict]:
        """统一调用 Microsoft Graph API。"""
        self.last_graph_error = ""

        # Token 快过期时先刷新；刷新失败就直接返回，避免进入无限重试。
        if not self.access_token or time.time() >= self.token_expires_at - 60:
            if not self.refresh_token():
                self.last_graph_error = "Access Token 不可用，Refresh Token 刷新失败"
                logger.error(self.last_graph_error)
                return None

        url = endpoint if str(endpoint).startswith(("http://", "https://")) else f"{self.graph_base}{endpoint}"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json",
        }
        headers.update(kwargs.pop("headers", {}) or {})
        timeout = kwargs.pop("timeout", 30)
        refreshed_after_401 = False

        for attempt in range(self.MAX_RETRIES + 1):
            try:
                response = requests.request(
                    method, url, headers=headers, timeout=timeout, **kwargs
                )

                # Access token 过期/失效：刷新一次后重试一次。
                if response.status_code == 401 and not refreshed_after_401:
                    refreshed_after_401 = True
                    if self.refresh_token():
                        headers["Authorization"] = f"Bearer {self.access_token}"
                        continue

                # 限流或服务端临时错误。
                if response.status_code in self.RETRY_STATUS and attempt < self.MAX_RETRIES:
                    delay = self._retry_delay(response, attempt)
                    logger.warning(
                        "Graph 请求重试 %s/%s：HTTP %s，%.1f 秒后重试",
                        attempt + 1, self.MAX_RETRIES, response.status_code, delay
                    )
                    time.sleep(delay)
                    continue

                if response.status_code >= 400:
                    try:
                        payload = response.json()
                    except ValueError:
                        payload = {}
                    error_obj = payload.get("error") if isinstance(payload, dict) else None
                    if isinstance(error_obj, dict):
                        detail = f"{error_obj.get('code', 'GraphError')}: {error_obj.get('message', response.text[:500])}"
                    else:
                        detail = response.text[:500] or f"HTTP {response.status_code}"
                    self.last_graph_error = f"HTTP {response.status_code}: {detail}"
                    logger.error("Graph 请求失败：%s", self.last_graph_error)
                    return None

                if response.status_code == 204 or not response.content:
                    return None

                try:
                    return response.json()
                except ValueError:
                    self.last_graph_error = "Graph 返回的数据不是有效 JSON"
                    logger.error("Graph JSON 解析失败：%s", response.text[:500])
                    return None

            except requests.RequestException as exc:
                if attempt >= self.MAX_RETRIES:
                    self.last_graph_error = f"Graph 网络请求失败：{exc}"
                    logger.error(self.last_graph_error)
                    return None
                delay = min(2 ** attempt, 16)
                logger.warning(
                    "Graph 网络请求重试 %s/%s：%.1f 秒后重试：%s",
                    attempt + 1, self.MAX_RETRIES, delay, exc
                )
                time.sleep(delay)

        self.last_graph_error = "Graph 请求超过最大重试次数"
        logger.error(self.last_graph_error)
        return None

    def get_user_info(self) -> typing.Optional[typing.Dict]:
        return self._request(
            "GET",
            "/me",
            params={"$select": "id,displayName,mail,userPrincipalName"},
        )

    def get_drive_info(self) -> typing.Dict:
        result = self._request('GET', '/me/drive')
        return result or {}

    def list_files(self, path: str = "/") -> typing.List[typing.Dict]:
        if path == "/" or path == "":
            endpoint = "/me/drive/root/children"
        else:
            endpoint = f"/me/drive/root:/{self._encode_path(path)}:/children"
        params = {'$select': 'id,name,size,lastModifiedDateTime,file,folder,eTag'}
        result = self._request('GET', endpoint, params=params)
        if result is None:
            raise RuntimeError(f"无法读取云端目录: {path or '/'}")
        items = list(result.get('value', []))
        next_link = result.get('@odata.nextLink')
        while next_link:
            page = self._request('GET', next_link)
            if page is None:
                break
            items.extend(page.get('value', []))
            next_link = page.get('@odata.nextLink')
        return items

    def download_file(self, remote_path: str, local_path: str) -> bool:
        try:
            file_info = self._request('GET', f"/me/drive/root:/{self._encode_path(remote_path)}")
            if not file_info or 'file' not in file_info:
                return False
            download_url = file_info.get('@microsoft.graph.downloadUrl')
            if not download_url:
                return False
            target_dir = os.path.dirname(local_path) or "."
            os.makedirs(target_dir, exist_ok=True)
            part_path = local_path + ".part"
            for attempt in range(self.MAX_RETRIES + 1):
                try:
                    response = requests.get(download_url, stream=True, timeout=120)
                    if response.status_code in self.RETRY_STATUS and attempt < self.MAX_RETRIES:
                        time.sleep(self._retry_delay(response, attempt))
                        continue
                    response.raise_for_status()
                    with open(part_path, 'wb') as f:
                        for chunk in response.iter_content(chunk_size=1024 * 1024):
                            if chunk:
                                f.write(chunk)
                    os.replace(part_path, local_path)
                    logger.info(f"📥 下载: {remote_path}")
                    return True
                except (requests.RequestException, OSError) as exc:
                    if attempt >= self.MAX_RETRIES:
                        raise
                    time.sleep(min(2 ** attempt, 16))
                    logger.warning(f"下载重试 {attempt + 1}: {remote_path} - {exc}")
            return False
        except (GraphAPIError, requests.RequestException, OSError) as e:
            logger.error(f"下载失败: {remote_path}: {e}")
            return False
        finally:
            part_path = local_path + ".part"
            if os.path.exists(part_path):
                try:
                    os.remove(part_path)
                except OSError:
                    pass

    def upload_file(self, local_path: str, remote_path: str) -> bool:
        try:
            file_size = os.path.getsize(local_path)
            encoded_path = self._encode_path(remote_path)
            if not encoded_path:
                encoded_path = self._encode_path(os.path.basename(local_path))
            endpoint = f"/me/drive/root:/{encoded_path}:/content"
            if file_size > 4 * 1024 * 1024:
                return self._upload_large_file(local_path, remote_path)
            with open(local_path, 'rb') as f:
                content = f.read()
            result = self._request('PUT', endpoint, headers={'Content-Type': 'application/octet-stream'}, data=content)
            logger.info(f"📤 上传: {remote_path}")
            return result is not None
        except (GraphAPIError, OSError) as e:
            logger.error(f"上传失败: {remote_path}: {e}")
            return False

    def _upload_large_file(self, local_path: str, remote_path: str) -> bool:
        encoded_path = self._encode_path(remote_path) or self._encode_path(os.path.basename(local_path))
        session_endpoint = f"/me/drive/root:/{encoded_path}:/createUploadSession"
        session_data = {"item": {"@microsoft.graph.conflictBehavior": "replace"}}
        try:
            result = self._request('POST', session_endpoint, json=session_data)
            upload_url = result.get('uploadUrl') if result else None
            if not upload_url:
                return False
            file_size = os.path.getsize(local_path)
            chunk_size = 4 * 1024 * 1024
            with open(local_path, 'rb') as f:
                offset = 0
                while offset < file_size:
                    chunk = f.read(chunk_size)
                    if not chunk:
                        break
                    end = offset + len(chunk) - 1
                    headers = {'Content-Length': str(len(chunk)), 'Content-Range': f'bytes {offset}-{end}/{file_size}'}
                    sent = False
                    for attempt in range(self.MAX_RETRIES + 1):
                        try:
                            response = requests.put(upload_url, headers=headers, data=chunk, timeout=120)
                            if response.status_code in self.RETRY_STATUS and attempt < self.MAX_RETRIES:
                                time.sleep(self._retry_delay(response, attempt))
                                continue
                            if response.status_code not in (200, 201, 202):
                                raise GraphAPIError(response.status_code, f"分片上传失败: {response.text[:500]}", response=response)
                            sent = True
                            break
                        except requests.RequestException as exc:
                            if attempt >= self.MAX_RETRIES:
                                raise
                            time.sleep(min(2 ** attempt, 16))
                    if not sent:
                        return False
                    offset += len(chunk)
            return True
        except (GraphAPIError, OSError, requests.RequestException) as e:
            logger.error(f"分片上传错误: {remote_path}: {e}")
            return False


class SyncState:
    """同步状态管理：保存本地快速指纹，避免每次扫描都计算 MD5。"""

    def __init__(self, state_file: str = STATE_FILE):
        self.state_file = state_file
        self.state = self.load()

    def load(self) -> typing.Dict:
        if os.path.exists(self.state_file):
            try:
                with open(self.state_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return data if isinstance(data, dict) else {}
            except (OSError, json.JSONDecodeError) as exc:
                logger.warning(f"同步状态读取失败，将使用空状态: {exc}")
        return {}

    def save(self):
        os.makedirs(os.path.dirname(self.state_file) or ".", exist_ok=True)
        tmp_file = self.state_file + ".tmp"
        with open(tmp_file, 'w', encoding='utf-8') as f:
            json.dump(self.state, f, indent=2, ensure_ascii=False)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_file, self.state_file)
        if platform.system() != 'Windows':
            try:
                os.chmod(self.state_file, 0o600)
            except OSError:
                pass

    @staticmethod
    def _fingerprint(filepath: str) -> typing.Tuple[int, int]:
        st = os.stat(filepath)
        return int(st.st_size), int(st.st_mtime_ns)

    def get_file_md5(self, filepath: str) -> str:
        md5 = hashlib.md5()
        try:
            with open(filepath, 'rb') as f:
                for chunk in iter(lambda: f.read(1024 * 1024), b''):
                    md5.update(chunk)
            return md5.hexdigest()
        except OSError as exc:
            raise RuntimeError(f"无法读取文件计算 MD5: {filepath}: {exc}") from exc

    def is_changed(self, relative_path: str, local_path: str) -> bool:
        old = self.state.get(relative_path)
        if not old:
            return True
        try:
            size, mtime_ns = self._fingerprint(local_path)
        except OSError:
            return True
        if old.get('size') == size and old.get('mtime_ns') == mtime_ns:
            return False
        current_md5 = self.get_file_md5(local_path)
        return current_md5 != old.get('md5', '')

    def update(self, relative_path: str, local_path: str):
        size, mtime_ns = self._fingerprint(local_path)
        self.state[relative_path] = {
            'md5': self.get_file_md5(local_path),
            'size': size,
            'mtime_ns': mtime_ns,
            'mtime': os.path.getmtime(local_path),
            'synced_at': datetime.datetime.now().isoformat()
        }

    def remove(self, relative_path: str):
        self.state.pop(relative_path, None)


class OneDriveSync:
    """同步引擎"""
    
    def __init__(self, client: OneDriveClient, config: typing.Dict, progress_callback=None, stop_event=None):
        self.client = client
        self.config = config
        state_identity = "|".join([
            config.get('user_email', ''), config.get('client_id', ''),
            config.get('local_folder', ''), config.get('root_folder', '/')
        ])
        state_name = hashlib.sha256(state_identity.encode('utf-8')).hexdigest()[:20] + ".json"
        self.state = SyncState(os.path.join(APP_DATA, "sync_state", state_name))
        os.makedirs(os.path.dirname(self.state.state_file), exist_ok=True)
        self.progress_callback = progress_callback
        self.stop_event = stop_event or threading.Event()
        
        self.local_folder = config.get('local_folder', '')
        if not self.local_folder:
            self.local_folder = os.path.join(os.path.expanduser("~"), "OneDrive")
        
        self.root_folder = config.get('root_folder', '/').strip('/')
    
    def log(self, message: str):
        if self.progress_callback:
            self.progress_callback(message)
        logger.info(message)

    def remote_path(self, rel_path: str) -> str:
        rel_path = rel_path.replace('\\', '/')
        root = self.root_folder.strip('/')
        return '/'.join(part for part in (root, rel_path.strip('/')) if part)
    
    def get_remote_files(self, path: str = "") -> typing.Dict[str, typing.Dict]:
        files = {}
        
        def _fetch(remote_path: str, prefix: str = ""):
            items = self.client.list_files(remote_path)
            for item in items:
                name = item.get('name', '')
                rel_path = f"{prefix}{name}".strip('/')
                
                if 'folder' in item:
                    _fetch(f"{remote_path}/{name}".strip('/'), f"{rel_path}/")
                else:
                    mtime_str = item.get('lastModifiedDateTime', '')
                    files[rel_path] = {
                        'id': item.get('id'),
                        'size': item.get('size', 0),
                        'mtime': mtime_str
                    }
        
        _fetch(self.root_folder)
        return files
    
    def get_local_files(self) -> typing.Dict[str, typing.Dict]:
        files = {}
        
        if not os.path.exists(self.local_folder):
            return files
        
        for root, dirs, filenames in os.walk(self.local_folder):
            dirs[:] = [d for d in dirs if not d.startswith('.')]
            
            for filename in filenames:
                if filename.startswith('.'):
                    continue
                
                filepath = os.path.join(root, filename)
                rel_path = os.path.relpath(filepath, self.local_folder)
                
                try:
                    files[rel_path] = {
                        'path': filepath,
                        'size': os.path.getsize(filepath),
                        'mtime': os.path.getmtime(filepath)
                    }
                except:
                    pass
        
        return files
    
    def sync_full(self) -> bool:
        self.log("="*50)
        self.log("🔄 完整同步...")
        
        os.makedirs(self.local_folder, exist_ok=True)
        
        local_files = self.get_local_files()
        remote_files = self.get_remote_files()
        
        self.log(f"   本地: {len(local_files)} 文件")
        self.log(f"   远程: {len(remote_files)} 文件")
        
        upload_count = 0
        download_count = 0
        
        for rel_path, local_info in local_files.items():
            if self.stop_event.is_set():
                self.log("同步已取消")
                self.state.save()
                return False
            remote_path = self.remote_path(rel_path)
            
            if rel_path in remote_files:
                if not self.state.is_changed(rel_path, local_info['path']):
                    continue
            
            if self.client.upload_file(local_info['path'], remote_path):
                self.state.update(rel_path, local_info['path'])
                upload_count += 1
                self.log(f"📤 上传: {rel_path}")
        
        for rel_path, remote_info in remote_files.items():
            if self.stop_event.is_set():
                self.log("同步已取消")
                self.state.save()
                return False
            local_path = os.path.join(self.local_folder, rel_path)
            
            if rel_path not in local_files:
                if self.client.download_file(
                    self.remote_path(rel_path),
                    local_path
                ):
                    if os.path.exists(local_path):
                        self.state.update(rel_path, local_path)
                    download_count += 1
                    self.log(f"📥 下载: {rel_path}")
        
        self.state.save()
        
        self.log(f"✅ 完成! 上传: {upload_count}, 下载: {download_count}")
        return True
    
    def sync_upload(self) -> bool:
        self.log("📤 仅上传模式...")
        
        local_files = self.get_local_files()
        remote_files = self.get_remote_files()
        
        upload_count = 0
        
        for rel_path, local_info in local_files.items():
            if self.stop_event.is_set():
                self.log("同步已取消")
                self.state.save()
                return False
            remote_path = self.remote_path(rel_path)
            
            if rel_path in remote_files:
                if not self.state.is_changed(rel_path, local_info['path']):
                    continue
            
            if self.client.upload_file(local_info['path'], remote_path):
                self.state.update(rel_path, local_info['path'])
                upload_count += 1
                self.log(f"📤 上传: {rel_path}")
        
        self.state.save()
        
        self.log(f"✅ 上传完成: {upload_count} 文件")
        return True
    
    def sync_download(self) -> bool:
        self.log("📥 仅下载模式...")
        
        local_files = self.get_local_files()
        remote_files = self.get_remote_files()
        
        download_count = 0
        
        for rel_path, remote_info in remote_files.items():
            if self.stop_event.is_set():
                self.log("同步已取消")
                self.state.save()
                return False
            local_path = os.path.join(self.local_folder, rel_path)
            
            if rel_path not in local_files:
                if self.client.download_file(
                    self.remote_path(rel_path),
                    local_path
                ):
                    if os.path.exists(local_path):
                        self.state.update(rel_path, local_path)
                    download_count += 1
                    self.log(f"📥 下载: {rel_path}")
        
        self.state.save()
        
        self.log(f"✅ 下载完成: {download_count} 文件")
        return True


# ============== 配置管理 ==============

def _protect_secret(value: str) -> str:
    """Windows 下用 DPAPI 保护 refresh_token/client_secret；其他系统保持兼容。"""
    if not value or not platform.system() == 'Windows' or value.startswith('dpapi:'):
        return value
    try:
        class DATA_BLOB(ctypes.Structure):
            _fields_ = [('cbData', ctypes.wintypes.DWORD), ('pbData', ctypes.POINTER(ctypes.c_byte))]
        raw = value.encode('utf-8')
        in_blob = DATA_BLOB(len(raw), ctypes.cast(ctypes.create_string_buffer(raw), ctypes.POINTER(ctypes.c_byte)))
        out_blob = DATA_BLOB()
        crypt = ctypes.windll.crypt32.CryptProtectData
        crypt.argtypes = [ctypes.POINTER(DATA_BLOB), ctypes.wintypes.LPCWSTR, ctypes.POINTER(DATA_BLOB), ctypes.wintypes.LPCWSTR, ctypes.c_void_p, ctypes.wintypes.DWORD, ctypes.POINTER(DATA_BLOB)]
        crypt.restype = ctypes.wintypes.BOOL
        if not crypt(ctypes.byref(in_blob), "OneDriveSync", None, None, None, 0, ctypes.byref(out_blob)):
            return value
        try:
            data = ctypes.string_at(out_blob.pbData, out_blob.cbData)
            return 'dpapi:' + base64.b64encode(data).decode('ascii')
        finally:
            ctypes.windll.kernel32.LocalFree(out_blob.pbData)
    except Exception as exc:
        logger.warning(f"DPAPI 加密失败，将保持兼容明文存储: {exc}")
        return value


def _unprotect_secret(value: str) -> str:
    if not isinstance(value, str) or not value.startswith('dpapi:') or platform.system() != 'Windows':
        return value
    try:
        class DATA_BLOB(ctypes.Structure):
            _fields_ = [('cbData', ctypes.wintypes.DWORD), ('pbData', ctypes.POINTER(ctypes.c_byte))]
        raw = base64.b64decode(value[6:])
        in_buf = ctypes.create_string_buffer(raw)
        in_blob = DATA_BLOB(len(raw), ctypes.cast(in_buf, ctypes.POINTER(ctypes.c_byte)))
        out_blob = DATA_BLOB()
        crypt = ctypes.windll.crypt32.CryptUnprotectData
        crypt.argtypes = [ctypes.POINTER(DATA_BLOB), ctypes.POINTER(ctypes.wintypes.LPWSTR), ctypes.POINTER(DATA_BLOB), ctypes.wintypes.LPCWSTR, ctypes.c_void_p, ctypes.wintypes.DWORD, ctypes.POINTER(DATA_BLOB)]
        crypt.restype = ctypes.wintypes.BOOL
        if not crypt(ctypes.byref(in_blob), None, None, None, None, 0, ctypes.byref(out_blob)):
            return value
        try:
            return ctypes.string_at(out_blob.pbData, out_blob.cbData).decode('utf-8')
        finally:
            ctypes.windll.kernel32.LocalFree(out_blob.pbData)
    except Exception as exc:
        logger.warning(f"DPAPI 解密失败: {exc}")
        return ''


def load_config() -> typing.Dict:
    default_config = {"accounts": [], "current_account": 0}
    if not os.path.exists(CONFIG_FILE):
        return default_config
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return default_config
        data.setdefault('accounts', [])
        data.setdefault('current_account', 0)
        for account in data['accounts']:
            for key in ('client_secret', 'refresh_token'):
                if key in account:
                    account[key] = _unprotect_secret(account[key])
        return data
    except (OSError, json.JSONDecodeError) as exc:
        logger.error(f"配置加载失败: {exc}")
        return default_config


def save_config(config: typing.Dict):
    os.makedirs(APP_DATA, exist_ok=True)
    safe_config = json.loads(json.dumps(config, ensure_ascii=False))
    for account in safe_config.get('accounts', []):
        for key in ('client_secret', 'refresh_token'):
            if key in account:
                account[key] = _protect_secret(account[key])
    tmp_file = CONFIG_FILE + ".tmp"
    try:
        with open(tmp_file, 'w', encoding='utf-8') as f:
            json.dump(safe_config, f, indent=2, ensure_ascii=False)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_file, CONFIG_FILE)
        if platform.system() != 'Windows':
            try:
                os.chmod(CONFIG_FILE, 0o600)
            except OSError:
                pass
    except OSError as exc:
        logger.error(f"配置保存失败: {exc}")
        try:
            if os.path.exists(tmp_file):
                os.remove(tmp_file)
        except OSError:
            pass
        raise


# ============== 自定义组件 ==============

class ModernButton(ctk.CTkButton):
    """现代化按钮"""
    def __init__(self, master, **kwargs):
        kwargs.setdefault('corner_radius', 8)
        kwargs.setdefault('border_width', 0)
        kwargs.setdefault('font', ctk.CTkFont(size=13, weight="bold"))
        kwargs.setdefault('height', 40)
        
        style = kwargs.pop('style', 'primary')
        
        if style == 'primary':
            kwargs.setdefault('fg_color', Theme.ACCENT)
            kwargs.setdefault('hover_color', Theme.ACCENT_HOVER)
            kwargs.setdefault('text_color', "#FFFFFF")
        elif style == 'secondary':
            kwargs.setdefault('fg_color', Theme.BG_CARD)
            kwargs.setdefault('hover_color', "#3D3D3D")
            kwargs.setdefault('text_color', Theme.TEXT_WHITE)
        elif style == 'success':
            kwargs.setdefault('fg_color', Theme.SUCCESS)
            kwargs.setdefault('hover_color', '#158a15')
            kwargs.setdefault('text_color', "#FFFFFF")
        elif style == 'warning':
            kwargs.setdefault('fg_color', Theme.WARNING)
            kwargs.setdefault('hover_color', '#E07B00')
            kwargs.setdefault('text_color', "#FFFFFF")
        elif style == 'danger':
            kwargs.setdefault('fg_color', Theme.ERROR)
            kwargs.setdefault('hover_color', '#c50f1f')
            kwargs.setdefault('text_color', "#FFFFFF")
        
        super().__init__(master, **kwargs)


class ModernCard(ctk.CTkFrame):
    """卡片组件"""
    def __init__(self, master, **kwargs):
        kwargs.setdefault('corner_radius', 12)
        kwargs.setdefault('fg_color', Theme.BG_CARD)
        kwargs.setdefault('border_width', 1)
        kwargs.setdefault('border_color', Theme.BORDER)
        super().__init__(master, **kwargs)


class ModernEntry(ctk.CTkEntry):
    """输入框组件"""
    def __init__(self, master, **kwargs):
        kwargs.setdefault('corner_radius', 6)
        kwargs.setdefault('border_width', 1)
        kwargs.setdefault('fg_color', "#1A1A1A")
        kwargs.setdefault('border_color', Theme.BORDER)
        kwargs.setdefault('font', ctk.CTkFont(size=13))
        super().__init__(master, **kwargs)


# ============== 对话框类 ==============

class AccountSelectDialog(ctk.CTkToplevel):
    """稳定的账号选择对话框。

    不使用 wait_window() 阻塞主窗口；通过 on_result 回调把结果交给主窗口。
    这样即使窗口创建、绘制或关闭过程中出现异常，也不会把主程序锁死。
    """
    def __init__(self, parent, accounts, mode="select", current_idx=-1, on_result=None):
        super().__init__(parent)
        self.parent = parent
        self.accounts = [a for a in (accounts or []) if isinstance(a, dict)]
        self.mode = mode
        self.current_idx = current_idx
        self.selected_idx = (
            current_idx
            if mode == "select" and 0 <= current_idx < len(self.accounts)
            else None
        )
        self.on_result = on_result
        self.row_buttons = []
        self._result_sent = False

        self.title("切换账号" if mode == "select" else "删除账号")
        # 固定小窗口，避免账号较少时窗口显得过大。
        self.geometry("150x300")
        self.resizable(False, False)
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", self.cancel)

        try:
            self.create_widgets()
            self.center_window(parent)
            self.update_idletasks()
            # 先完成界面创建，再设置 grab，避免创建控件异常时遗留全局 grab。
            self.after(20, self._activate_modal)
        except Exception:
            logger.exception("创建账号选择窗口失败")
            self._safe_release_grab()
            try:
                self.destroy()
            except Exception:
                pass
            raise

    def _activate_modal(self):
        try:
            if not self.winfo_exists():
                return
            self.lift()
            self.focus_force()
            self.grab_set()
        except Exception:
            logger.exception("激活账号选择窗口失败")

    def _safe_release_grab(self):
        try:
            self.grab_release()
        except Exception:
            pass

    def center_window(self, parent):
        self.update_idletasks()
        width = self.winfo_width() or 150
        height = self.winfo_height() or 300
        x = parent.winfo_x() + max(0, (parent.winfo_width() - width) // 2)
        y = parent.winfo_y() + max(0, (parent.winfo_height() - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    @staticmethod
    def _account_text(account):
        name = str(account.get("name") or "未命名账号").strip()
        email = str(account.get("user_email") or "未获取邮箱").strip()
        return name, email

    def create_widgets(self):
        self.configure(fg_color=Theme.BG_MAIN)

        title = "切换账号" if self.mode == "select" else "删除账号"
        ctk.CTkLabel(
            self,
            text=title,
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(pady=(12, 8))

        self.account_listbox = ctk.CTkScrollableFrame(
            self,
            fg_color="transparent",
            corner_radius=0,
            height=170,
        )
        self.account_listbox.pack(fill="x", padx=8, pady=(0, 6))
        self.account_listbox.grid_columnconfigure(0, weight=1)

        for idx, acc in enumerate(self.accounts):
            name, email = self._account_text(acc)
            # 仅显示账号名称和邮箱，当前账号通过背景色区分。
            text = f"{name}\n{email}"
            button = ctk.CTkButton(
                self.account_listbox,
                text=text,
                anchor="w",
                height=50,
                corner_radius=8,
                fg_color=Theme.ACCENT if idx == self.selected_idx else Theme.BG_CARD,
                hover_color=Theme.ACCENT_HOVER,
                border_width=1,
                border_color=Theme.ACCENT_HOVER if idx == self.selected_idx else Theme.BORDER,
                font=ctk.CTkFont(size=10),
                command=lambda i=idx: self.select_account(i),
            )
            button.pack(fill="x", padx=2, pady=3)
            self.row_buttons.append(button)

        if not self.accounts:
            ctk.CTkLabel(
                self.account_listbox,
                text="暂无账号",
                font=ctk.CTkFont(size=12),
                text_color=Theme.TEXT_GRAY,
            ).pack(expand=True, pady=28)

        button_bar = ctk.CTkFrame(self, fg_color="transparent")
        button_bar.pack(fill="x", padx=8, pady=(4, 10))

        self.confirm_button = ModernButton(
            button_bar,
            text="确认切换" if self.mode == "select" else "确认删除",
            style="primary" if self.mode == "select" else "danger",
            command=self.confirm,
            height=30,
        )
        self.confirm_button.pack(side="left", fill="x", expand=True, padx=4)

        ModernButton(
            button_bar,
            text="取消",
            style="secondary",
            command=self.cancel,
            height=30,
        ).pack(side="left", fill="x", expand=True, padx=4)

        if self.mode == "select" and self.accounts:
            self.confirm_button.focus_set()

    def select_account(self, idx):
        if not 0 <= idx < len(self.accounts):
            return
        self.selected_idx = idx
        for i, button in enumerate(self.row_buttons):
            selected = i == idx
            button.configure(
                fg_color=Theme.ACCENT if selected else Theme.BG_CARD,
                hover_color=Theme.ACCENT_HOVER,
                border_color=Theme.ACCENT_HOVER if selected else Theme.BORDER,
            )

    def _finish(self, result):
        if self._result_sent:
            return
        self._result_sent = True
        self.selected_idx = result
        callback = self.on_result
        self._safe_release_grab()
        try:
            self.destroy()
        except Exception:
            pass
        if callback:
            try:
                self.parent.after(0, callback, result)
            except Exception:
                logger.exception("返回账号选择结果失败")

    def confirm(self):
        if self.selected_idx is None:
            messagebox.showwarning("请选择账号", "请先点击一个账号。", parent=self)
            return
        if self.mode == "delete":
            account = self.accounts[self.selected_idx]
            name = str(account.get("name") or f"账号{self.selected_idx + 1}")
            if not messagebox.askyesno("确认删除", f"确定要删除账号“{name}”吗？", parent=self):
                return
        self._finish(self.selected_idx)

    def cancel(self):
        self._finish(None)


class AddAccountDialog(ctk.CTkToplevel):
    """添加账号对话框"""
    def __init__(self, parent, defaults=None):
        super().__init__(parent)
        self.title("添加 OneDrive 账号")
        self.geometry("500x690")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        self.result = None
        self.defaults = (defaults or {}).copy()
        
        self.create_widgets()
        self.center_window(parent)
    
    def center_window(self, parent):
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (500 // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (690 // 2)
        self.geometry(f"500x690+{x}+{y}")
    
    def create_widgets(self):
        self.configure(fg_color=Theme.BG_MAIN)
        
        # 标题
        ctk.CTkLabel(
            self,
            text="添加 OneDrive 账号",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(pady=15)
        
        # 表单
        form_frame = ctk.CTkScrollableFrame(self)
        form_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # 账号名称
        ctk.CTkLabel(form_frame, text="账号名称:").pack(anchor="w", pady=(10, 2))
        self.name_entry = ModernEntry(form_frame, placeholder_text="例如: 工作账号")
        self.name_entry.pack(fill="x", pady=(0, 10))
        
        # 区域选择
        ctk.CTkLabel(form_frame, text="区域:").pack(anchor="w", pady=(10, 2))
        self.region_var = ctk.StringVar(value=self.defaults.get("region", "global"))
        region_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        region_frame.pack(fill="x", pady=(0, 10))
        
        regions = [("全球版", "global"), ("中国版", "cn"), ("德国版", "de"), ("美国版", "us")]
        for text, val in regions:
            ctk.CTkRadioButton(
                region_frame,
                text=text,
                variable=self.region_var,
                value=val
            ).pack(side="left", padx=5)
        
        # Client ID
        ctk.CTkLabel(form_frame, text="Client ID:").pack(anchor="w", pady=(10, 2))
        self.client_id_entry = ModernEntry(form_frame, placeholder_text="应用程序 ID")
        if self.defaults.get("client_id"):
            self.client_id_entry.insert(0, self.defaults["client_id"])
        self.client_id_entry.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(
            form_frame,
            text="所需权限：User.Read、Files.ReadWrite.All、offline_access",
            text_color=Theme.TEXT_GRAY,
        ).pack(anchor="w", pady=(0, 6))
        
        # Client Secret
        ctk.CTkLabel(form_frame, text="Client Secret（公共客户端无需填写）:").pack(anchor="w", pady=(10, 2))
        self.client_secret_entry = ModernEntry(form_frame, placeholder_text="公共客户端请留空", show="•")
        if self.defaults.get("client_secret"):
            self.client_secret_entry.insert(0, self.defaults["client_secret"])
        self.client_secret_entry.pack(fill="x", pady=(0, 10))
        
        # Tenant ID
        ctk.CTkLabel(form_frame, text="Tenant ID:").pack(anchor="w", pady=(10, 2))
        self.tenant_id_entry = ModernEntry(form_frame, placeholder_text="默认: common")
        self.tenant_id_entry.insert(0, self.defaults.get("tenant_id", "common") or "common")
        self.tenant_id_entry.pack(fill="x", pady=(0, 10))
        
        # 本地文件夹
        ctk.CTkLabel(form_frame, text="本地同步文件夹:").pack(anchor="w", pady=(10, 2))
        folder_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        folder_frame.pack(fill="x", pady=(0, 10))
        
        self.local_folder_entry = ModernEntry(folder_frame, placeholder_text="选择本地文件夹")
        self.local_folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        ModernButton(
            folder_frame,
            text="浏览",
            style="secondary",
            command=self.select_local_folder,
            height=30,
            width=80
        ).pack(side="right")

        # 按钮
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=15)
        
        self.auth_action_button = ModernButton(
            btn_frame,
            text="登录并自动授权",
            style="primary",
            command=self.save_account,
            height=36
        )
        self.auth_action_button.pack(side="left", fill="x", expand=True, padx=5)
        
        ModernButton(
            btn_frame,
            text="取消",
            style="secondary",
            command=self.cancel,
            height=36
        ).pack(side="left", fill="x", expand=True, padx=5)
    
    def select_local_folder(self):
        folder = filedialog.askdirectory(title="选择本地同步文件夹")
        if folder:
            self.local_folder_entry.delete(0, tk.END)
            self.local_folder_entry.insert(0, folder)
    
    def _start_local_callback(self, account):
        """启动一次性 localhost OAuth 回调监听器。"""
        redirect = urllib.parse.urlparse(account['redirect_uri'])
        host = redirect.hostname or 'localhost'
        port = redirect.port or 80
        expected_path = redirect.path or '/'
        result_queue = queue.Queue(maxsize=1)
        dialog = self

        class CallbackHandler(http.server.BaseHTTPRequestHandler):
            def do_GET(self):
                parsed = urllib.parse.urlparse(self.path)
                if parsed.path != expected_path:
                    self.send_response(404)
                    self.end_headers()
                    return
                params = urllib.parse.parse_qs(parsed.query)
                result = {
                    'code': (params.get('code') or [''])[0],
                    'state': (params.get('state') or [''])[0],
                    'error': (params.get('error_description') or params.get('error') or [''])[0],
                }
                try:
                    result_queue.put_nowait(result)
                except queue.Full:
                    pass
                body = ("<html><meta charset='utf-8'><body style='font-family:Segoe UI;padding:40px'>"
                        "<h2>OneDrive 授权已返回</h2><p>可以关闭此浏览器窗口并返回同步工具。</p>"
                        "</body></html>").encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def log_message(self, format, *args):
                return

        class CallbackServer(socketserver.TCPServer):
            allow_reuse_address = True

        try:
            server = CallbackServer((host, port), CallbackHandler)
        except OSError as e:
            raise RuntimeError(f"无法监听 {host}:{port}，端口可能已被占用: {e}") from e

        server.timeout = 1
        self.callback_server = server
        self.callback_result_queue = result_queue

        def wait_callback():
            deadline = time.time() + 180
            try:
                while time.time() < deadline and self.winfo_exists():
                    server.handle_request()
                    try:
                        result = result_queue.get_nowait()
                    except queue.Empty:
                        continue
                    self.after(0, self._finish_local_authorization, account, result)
                    return
                if self.winfo_exists():
                    self.after(0, self._authorization_timeout)
            finally:
                server.server_close()

        threading.Thread(target=wait_callback, daemon=True).start()

    def _authorization_timeout(self):
        self.auth_action_button.configure(text="登录并自动授权", state="normal")
        messagebox.showwarning("授权超时", "180 秒内没有收到 localhost 授权回调，请重新登录。")

    def _finish_local_authorization(self, account, result):
        if result.get('error'):
            self.auth_action_button.configure(text="登录并自动授权", state="normal")
            messagebox.showerror("授权失败", result['error'])
            return
        if not result.get('code'):
            self.auth_action_button.configure(text="登录并自动授权", state="normal")
            messagebox.showerror("授权失败", "localhost 回调中没有 code 参数")
            return
        if result.get('state') != getattr(self, 'pending_state', None):
            self.auth_action_button.configure(text="登录并自动授权", state="normal")
            messagebox.showerror("授权失败", "OAuth state 不匹配，请重新授权")
            return

        self.auth_action_button.configure(text="正在获取账号信息...", state="disabled")
        client = OneDriveClient(account)

        def exchange_worker():
            try:
                ok = client.exchange_code(result['code'])
                user_info = client.get_user_info() if ok else None
                # 即使 /me 权限不足，Token 已经成功时也结束授权流程，避免界面永久停留。
                self.after(0, self._complete_account_authorization, account, client, ok, user_info)
            except Exception as exc:
                logger.exception("授权后处理异常")
                client.last_auth_error = f"授权后处理异常: {exc}"
                self.after(0, self._complete_account_authorization, account, client, False, None)

        threading.Thread(target=exchange_worker, daemon=True).start()

    def _complete_account_authorization(self, account, client, ok, user_info):
        if not ok:
            self.auth_action_button.configure(text="登录并自动授权", state="normal")
            error_detail = getattr(client, "last_auth_error", "授权码交换 Token 失败")
            messagebox.showerror("授权失败", "授权码交换 Token 失败\n\n" + error_detail)
            logger.error("授权失败详情: %s", error_detail)
            return
        if user_info:
            account['user_email'] = user_info.get('mail') or user_info.get('userPrincipalName', '')
        elif getattr(client, 'last_graph_error', ''):
            logger.warning("Token 已成功获取，但读取账号信息失败：%s", client.last_graph_error)
        account['refresh_token'] = client.config.get('refresh_token', '')
        self.result = account
        self.destroy()

    def save_account(self):
        name = self.name_entry.get().strip()
        client_id = self.client_id_entry.get().strip()
        client_secret = self.client_secret_entry.get().strip()
        if not name or not client_id:
            messagebox.showerror("错误", "请填写账号名称和 Client ID")
            return

        account = ACCOUNT_TEMPLATE.copy()
        account.update({
            "name": name,
            "region": self.region_var.get(),
            "client_id": client_id,
            "client_secret": client_secret,
            "tenant_id": self.tenant_id_entry.get().strip() or "common",
            "redirect_uri": "http://localhost:53682/callback",
            "local_folder": self.local_folder_entry.get().strip()
        })
        try:
            self._start_local_callback(account)
            client = OneDriveClient(account)
            auth_url, state = client.get_authorization_url()
            self.pending_state = state
            self.auth_action_button.configure(text="等待浏览器授权...", state="disabled")
            opened = webbrowser.open(auth_url, new=2)
            if not opened:
                self.clipboard_clear()
                self.clipboard_append(auth_url)
                self.update()
                messagebox.showwarning("浏览器未打开", "授权地址已复制到剪贴板，请粘贴到浏览器打开。")
        except Exception as e:
            logger.exception("启动 localhost 授权失败")
            self.auth_action_button.configure(text="登录并自动授权", state="normal")
            messagebox.showerror("无法开始授权", str(e))

    def cancel(self):
        self.result = None
        server = getattr(self, 'callback_server', None)
        if server:
            try:
                server.server_close()
            except Exception:
                pass
        self.destroy()


class EditPathsDialog(ctk.CTkToplevel):
    """编辑路径对话框"""
    def __init__(self, parent, account):
        super().__init__(parent)
        self.title("修改同步路径")
        self.geometry("450x300")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        self.account = account.copy()
        self.result = None
        
        self.create_widgets()
        self.center_window(parent)
    
    def center_window(self, parent):
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (450 // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (300 // 2)
        self.geometry(f"450x300+{x}+{y}")
    
    def create_widgets(self):
        self.configure(fg_color=Theme.BG_MAIN)
        
        # 标题
        ctk.CTkLabel(
            self,
            text="修改同步路径",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(pady=15)
        
        # 表单
        form_frame = ctk.CTkFrame(self, fg_color="transparent")
        form_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # 远程根文件夹
        ctk.CTkLabel(form_frame, text="远程根文件夹:").pack(anchor="w", pady=(10, 2))
        self.root_folder_entry = ModernEntry(
            form_frame,
            placeholder_text="/",
            textvariable=tk.StringVar(value=self.account.get('root_folder', '/'))
        )
        self.root_folder_entry.pack(fill="x", pady=(0, 10))
        
        # 本地文件夹
        ctk.CTkLabel(form_frame, text="本地同步文件夹:").pack(anchor="w", pady=(10, 2))
        folder_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        folder_frame.pack(fill="x", pady=(0, 10))
        
        self.local_folder_entry = ModernEntry(
            folder_frame,
            textvariable=tk.StringVar(value=self.account.get('local_folder', ''))
        )
        self.local_folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        ModernButton(
            folder_frame,
            text="浏览",
            style="secondary",
            command=self.select_local_folder,
            height=30,
            width=80
        ).pack(side="right")
        
        # 按钮
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=15)
        
        ModernButton(
            btn_frame,
            text="保存",
            style="primary",
            command=self.save_paths,
            height=36
        ).pack(side="left", fill="x", expand=True, padx=5)
        
        ModernButton(
            btn_frame,
            text="取消",
            style="secondary",
            command=self.cancel,
            height=36
        ).pack(side="left", fill="x", expand=True, padx=5)
    
    def select_local_folder(self):
        folder = filedialog.askdirectory(title="选择本地同步文件夹")
        if folder:
            self.local_folder_entry.delete(0, tk.END)
            self.local_folder_entry.insert(0, folder)
    
    def save_paths(self):
        self.account['root_folder'] = self.root_folder_entry.get().strip() or "/"
        self.account['local_folder'] = self.local_folder_entry.get().strip()
        self.result = self.account
        self.destroy()
    
    def cancel(self):
        self.result = None
        self.destroy()


class AuthDialog(ctk.CTkToplevel):
    """重新授权对话框：公共客户端使用 localhost 回调自动完成授权。"""
    def __init__(self, parent, account):
        super().__init__(parent)
        self.title("重新授权")
        self.geometry("450x230")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.account = account
        self.success = False
        self.callback_server = None
        self.pending_state = None

        self.create_widgets()
        self.center_window(parent)

    def center_window(self, parent):
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (450 // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (230 // 2)
        self.geometry(f"450x230+{x}+{y}")

    def create_widgets(self):
        self.configure(fg_color=Theme.BG_MAIN)

        ctk.CTkLabel(
            self,
            text="需要重新授权",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(pady=15)

        ctk.CTkLabel(
            self,
            text="点击下方按钮后将在浏览器完成 Microsoft 登录，授权成功后自动返回。",
            text_color=Theme.TEXT_GRAY,
            wraplength=360
        ).pack(pady=8)

        self.auth_button = ModernButton(
            self,
            text="打开授权页面",
            style="primary",
            command=self.authorize,
            height=40
        )
        self.auth_button.pack(fill="x", padx=40, pady=20)

        ModernButton(
            self,
            text="取消",
            style="secondary",
            command=self.cancel,
            height=36
        ).pack(fill="x", padx=40, pady=(0, 10))

    def _start_local_callback(self):
        redirect = urllib.parse.urlparse(self.account.get('redirect_uri') or '')
        host = redirect.hostname or 'localhost'
        port = redirect.port or 80
        expected_path = redirect.path or '/'
        result_queue = queue.Queue(maxsize=1)
        dialog = self

        class CallbackHandler(http.server.BaseHTTPRequestHandler):
            def do_GET(self):
                parsed = urllib.parse.urlparse(self.path)
                if parsed.path != expected_path:
                    self.send_response(404)
                    self.end_headers()
                    return

                params = urllib.parse.parse_qs(parsed.query)
                result = {
                    'code': (params.get('code') or [''])[0],
                    'state': (params.get('state') or [''])[0],
                    'error': (params.get('error_description') or params.get('error') or [''])[0],
                }
                try:
                    result_queue.put_nowait(result)
                except queue.Full:
                    pass

                body = (
                    "<html><meta charset='utf-8'><body style='font-family:Segoe UI;padding:40px'>"
                    "<h2>OneDrive 授权已返回</h2><p>可以关闭此浏览器窗口并返回同步工具。</p>"
                    "</body></html>"
                ).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def log_message(self, format, *args):
                return

        class CallbackServer(socketserver.TCPServer):
            allow_reuse_address = True

        try:
            server = CallbackServer((host, port), CallbackHandler)
        except OSError as e:
            raise RuntimeError(f"无法监听 {host}:{port}，端口可能已被占用: {e}") from e

        server.timeout = 1
        self.callback_server = server

        def wait_callback():
            deadline = time.time() + 180
            try:
                while time.time() < deadline and self.winfo_exists():
                    server.handle_request()
                    try:
                        result = result_queue.get_nowait()
                    except queue.Empty:
                        continue
                    self.after(0, self._finish_authorization, result)
                    return
                if self.winfo_exists():
                    self.after(0, self._authorization_timeout)
            finally:
                try:
                    server.server_close()
                except Exception:
                    pass

        threading.Thread(target=wait_callback, daemon=True).start()

    def _authorization_timeout(self):
        self.auth_button.configure(text="打开授权页面", state="normal")
        messagebox.showwarning("授权超时", "180 秒内没有收到 localhost 授权回调，请重新授权。")

    def authorize(self):
        try:
            client = OneDriveClient(self.account)
            auth_url, state = client.get_authorization_url()
            self.pending_state = state
            self._start_local_callback()
            self.auth_button.configure(text="等待浏览器授权...", state="disabled")
            opened = webbrowser.open(auth_url, new=2)
            if not opened:
                self.clipboard_clear()
                self.clipboard_append(auth_url)
                self.update()
                messagebox.showwarning("浏览器未打开", "授权地址已复制到剪贴板，请粘贴到浏览器打开。")
        except Exception as e:
            logger.exception("重新授权启动失败")
            self.auth_button.configure(text="打开授权页面", state="normal")
            messagebox.showerror("无法开始授权", str(e))

    def _finish_authorization(self, result):
        if result.get('error'):
            self.auth_button.configure(text="打开授权页面", state="normal")
            messagebox.showerror("授权失败", result['error'])
            return
        if not result.get('code'):
            self.auth_button.configure(text="打开授权页面", state="normal")
            messagebox.showerror("授权失败", "localhost 回调中没有 code 参数")
            return
        if result.get('state') != self.pending_state:
            self.auth_button.configure(text="打开授权页面", state="normal")
            messagebox.showerror("授权失败", "OAuth state 不匹配，请重新授权")
            return

        self.auth_button.configure(text="正在获取账号信息...", state="disabled")
        client = OneDriveClient(self.account)

        def exchange_worker():
            try:
                ok = client.exchange_code(result['code'])
                user_info = client.get_user_info() if ok else None
                self.after(0, self._complete_authorization, client, ok, user_info)
            except Exception as exc:
                logger.exception("重新授权后处理异常")
                client.last_auth_error = f"重新授权后处理异常: {exc}"
                self.after(0, self._complete_authorization, client, False, None)

        threading.Thread(target=exchange_worker, daemon=True).start()

    def _complete_authorization(self, client, ok, user_info):
        if not ok:
            self.auth_button.configure(text="打开授权页面", state="normal")
            error_detail = getattr(client, "last_auth_error", "授权码交换 Token 失败")
            messagebox.showerror("授权失败", "授权码交换 Token 失败\n\n" + error_detail)
            logger.error("重新授权失败详情: %s", error_detail)
            return

        self.account['refresh_token'] = client.config.get('refresh_token', '')
        if user_info:
            self.account['user_email'] = user_info.get('mail') or user_info.get('userPrincipalName', '')
        elif getattr(client, 'last_graph_error', ''):
            logger.warning("重新授权成功，但读取账号信息失败：%s", client.last_graph_error)
        self.success = True
        self.destroy()

    def cancel(self):
        self.success = False
        server = getattr(self, 'callback_server', None)
        if server:
            try:
                server.server_close()
            except Exception:
                pass
        self.destroy()


class CloudBrowserDialog(ctk.CTkToplevel):
    """云端文件浏览器"""
    def __init__(self, parent, client):
        super().__init__(parent)
        self.title("云端文件浏览")
        self.geometry("700x500")
        self.transient(parent)
        self.grab_set()
        
        self.client = client
        self.current_path = ""
        
        self.create_widgets()
        self.center_window(parent)
        self.load_files_async()
    
    def center_window(self, parent):
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (700 // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (500 // 2)
        self.geometry(f"700x500+{x}+{y}")
    
    def create_widgets(self):
        self.configure(fg_color=Theme.BG_MAIN)
        
        # 路径导航
        path_frame = ctk.CTkFrame(self, fg_color=Theme.BG_CARD)
        path_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        self.path_label = ctk.CTkLabel(
            path_frame,
            text="/",
            font=ctk.CTkFont(size=13)
        )
        self.path_label.pack(side="left", padx=15, pady=10)
        
        # 文件列表
        self.files_frame = ctk.CTkScrollableFrame(self)
        self.files_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # 按钮
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=10)
        
        ModernButton(
            btn_frame,
            text="返回上一级",
            style="secondary",
            command=self.go_back,
            height=36
        ).pack(side="left", padx=5)
        
        ModernButton(
            btn_frame,
            text="关闭",
            style="primary",
            command=self.destroy,
            height=36
        ).pack(side="right", padx=5)
    
    def _clear_files(self):
        for child in self.files_frame.winfo_children():
            child.destroy()

    def load_files_async(self):
        """网络请求放到后台线程，避免 Tk 主线程卡死。"""
        self._clear_files()
        self.path_label.configure(text=f"/{self.current_path}" if self.current_path else "/")
        ctk.CTkLabel(self.files_frame, text="正在加载...", text_color=Theme.TEXT_GRAY).pack(pady=20)
        path = self.current_path

        def worker():
            try:
                files = self.client.list_files(path)
                self.after(0, self._render_files, path, files, None)
            except Exception as exc:
                self.after(0, self._render_files, path, [], exc)

        threading.Thread(target=worker, daemon=True).start()

    def _render_files(self, path, files, error):
        if not self.winfo_exists() or path != self.current_path:
            return
        self._clear_files()
        if error:
            ctk.CTkLabel(
                self.files_frame,
                text=f"加载失败: {error}",
                text_color=Theme.ERROR
            ).pack(pady=20)
            return
        if not files:
            ctk.CTkLabel(
                self.files_frame,
                text="此文件夹为空",
                text_color=Theme.TEXT_GRAY
            ).pack(pady=20)
            return

        for item in files:
            name = item.get('name', '')
            is_folder = 'folder' in item
            size = item.get('size', 0)
            mtime = item.get('lastModifiedDateTime', '')
            frame = ctk.CTkFrame(self.files_frame, fg_color=Theme.BG_CARD, corner_radius=8)
            frame.pack(fill="x", padx=5, pady=3)
            if is_folder:
                frame.bind("<Button-1>", lambda e, p=name: self.enter_folder(p))
                ctk.CTkLabel(frame, text="📁", font=ctk.CTkFont(size=16)).pack(side="left", padx=(10, 5), pady=8)
                ctk.CTkLabel(frame, text=name, font=ctk.CTkFont(size=13, weight="bold")).pack(side="left", anchor="center")
                ctk.CTkLabel(frame, text=mtime[:10] if mtime else "", font=ctk.CTkFont(size=11), text_color=Theme.TEXT_GRAY).pack(side="right", padx=10, pady=8)
            else:
                ctk.CTkLabel(frame, text="📄", font=ctk.CTkFont(size=16)).pack(side="left", padx=(10, 5), pady=8)
                ctk.CTkLabel(frame, text=name, font=ctk.CTkFont(size=13)).pack(side="left", anchor="center")
                size_str = f"{size/1024/1024:.2f} MB" if size > 1024*1024 else f"{size/1024:.0f} KB"
                ctk.CTkLabel(frame, text=size_str, font=ctk.CTkFont(size=11), text_color=Theme.TEXT_GRAY).pack(side="right", padx=10, pady=8)

    def enter_folder(self, folder_name):
        if self.current_path:
            self.current_path = f"{self.current_path}/{folder_name}"
        else:
            self.current_path = folder_name
        self.load_files_async()
    
    def go_back(self):
        if "/" in self.current_path:
            self.current_path = self.current_path.rsplit("/", 1)[0]
        else:
            self.current_path = ""
        self.load_files_async()


# ============== GUI 主窗口 ==============

class OneDriveGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("OneDrive 同步工具")
        self.geometry("600x580")
        self.resizable(False, False)
        
        # 配置
        self.config = load_config()
        self.accounts = self.config.get('accounts', [])
        self.current_idx = self._normalize_current_index(self.config.get('current_account', 0))
        self.config['current_account'] = self.current_idx
        self.config['accounts'] = self.accounts
        
        # 同步状态
        self.is_syncing = False
        self.sync_thread = None
        self.stop_event = threading.Event()
        self.log_queue = queue.Queue()
        self.cloud_dialog = None
        self.account_dialog = None
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # 创建界面
        self.create_widgets()
        self.after(100, self.process_log_queue)
        self.refresh_account_list()
        self.center_window()
    
    def on_close(self):
        """安全关闭主窗口，避免 modal/grab 子窗口导致无法退出。"""
        try:
            dialog = getattr(self, "account_dialog", None)
            if dialog is not None and dialog.winfo_exists():
                try:
                    dialog.cancel()
                except Exception:
                    dialog.destroy()
        except Exception:
            pass
        self.close_cloud_dialog()
        try:
            self.stop_event.set()
        except Exception:
            pass
        self.destroy()

    def center_window(self):
        """窗口居中"""
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (self.winfo_width() // 2)
        y = (self.winfo_screenheight() // 2) - (self.winfo_height() // 2)
        self.geometry(f'{self.winfo_width()}x{self.winfo_height()}+{x}+{y}')
    
    def create_widgets(self):
        """创建界面组件"""
        self.configure(fg_color=Theme.BG_MAIN)
        
        # ===== 新增: 顶部蓝色标题栏 =====
        title_bar = ctk.CTkFrame(self, fg_color=Theme.ACCENT, height=60, corner_radius=0)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)
        
        title_inner = ctk.CTkFrame(title_bar, fg_color="transparent")
        title_inner.place(relx=0.5, rely=0.5, anchor="center")
        
        ctk.CTkLabel(
            title_inner,
            text="☁️",
            font=ctk.CTkFont(size=28)
        ).pack(side="left", padx=(0, 8))
        
        ctk.CTkLabel(
            title_inner,
            text="OneDrive 同步工具",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color="#FFFFFF"
        ).pack(side="left")
        
        # 新增: 状态指示灯
        self.status_indicator = ctk.CTkLabel(
            title_bar,
            text="●",
            font=ctk.CTkFont(size=16),
            text_color="#BBBBBB"
        )
        self.status_indicator.pack(side="right", padx=20)
        
        # 账号卡片
        account_card = ModernCard(self)
        account_card.pack(fill="x", padx=20, pady=(20, 10))
        
        # 新增: 账号标题行
        account_header = ctk.CTkFrame(account_card, fg_color="transparent")
        account_header.pack(fill="x", padx=20, pady=(18, 10))
        
        ctk.CTkLabel(
            account_header,
            text="👤 当前账号",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=Theme.TEXT_WHITE
        ).pack(side="left")
        
        # 新增: 账号徽章
        self.account_badge = ctk.CTkLabel(
            account_header,
            text="",
            font=ctk.CTkFont(size=11),
            text_color=Theme.TEXT_MUTED
        )
        self.account_badge.pack(side="right")
        
        # 账号信息
        self.account_label = ctk.CTkLabel(
            account_card,
            text="未选择账号",
            font=ctk.CTkFont(size=13),
            text_color=Theme.TEXT_GRAY
        )
        self.account_label.pack(anchor="w", padx=20, pady=(0, 5))
        
        # 新增: 本地文件夹显示
        self.folder_label = ctk.CTkLabel(
            account_card,
            text="",
            font=ctk.CTkFont(size=11),
            text_color=Theme.TEXT_MUTED
        )
        self.folder_label.pack(anchor="w", padx=20, pady=(0, 15))
        
        # 账号按钮 - 样式优化
        account_btns = ctk.CTkFrame(account_card, fg_color="transparent")
        account_btns.pack(fill="x", padx=15, pady=(0, 15))
        
        ModernButton(
            account_btns,
            text="🔄 切换账号",
            style="primary",
            command=self.switch_account,
            height=38
        ).pack(side="left", fill="both", expand=True, padx=4)
        
        ModernButton(
            account_btns,
            text="➕ 添加账号",
            style="primary",
            command=self.add_account,
            height=38
        ).pack(side="left", fill="both", expand=True, padx=4)
        
        ModernButton(
            account_btns,
            text="🗑️删除账号",
            style="danger",
            command=self.delete_account,
            height=38
        ).pack(side="left", fill="both", expand=True, padx=4)
        
        # 同步操作卡片
        sync_card = ModernCard(self)
        sync_card.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(
            sync_card,
            text="⚡ 同步操作",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=Theme.TEXT_WHITE
        ).pack(anchor="w", padx=20, pady=(18, 15))
        
        # 新增: 主同步按钮 - 大卡片样式
        main_sync = ctk.CTkFrame(sync_card, fg_color=Theme.ACCENT, corner_radius=10)
        main_sync.pack(fill="x", padx=15, pady=(0, 12))
        
        main_inner = ctk.CTkFrame(main_sync, fg_color="transparent")
        main_inner.pack(fill="x", padx=20, pady=16)
        
        ctk.CTkLabel(
            main_inner,
            text="🔄",
            font=ctk.CTkFont(size=22)
        ).pack(side="left", padx=(0, 12))
        
        ctk.CTkLabel(
            main_inner,
            text="双端同步",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color="#FFFFFF"
        ).pack(side="left")
        
        ctk.CTkLabel(
            main_inner,
            text="上传 + 下载",
            font=ctk.CTkFont(size=11),
            text_color="#FFFFFF"
        ).pack(side="right")
        
        # 主按钮点击事件
        main_sync.bind("<Button-1>", lambda e: self.sync_full())
        main_inner.bind("<Button-1>", lambda e: self.sync_full())
        for widget in main_inner.winfo_children():
            widget.bind("<Button-1>", lambda e: self.sync_full())
        
        # 上下传并排 - 优化: 分开两个按钮，统一风格
        quick_sync = ctk.CTkFrame(sync_card, fg_color="transparent")
        quick_sync.pack(fill="x", padx=15, pady=(0, 12))
        
        # 上传按钮 - 绿色统一风格
        ModernButton(
            quick_sync,
            text="⬆️ 上传",
            style="success",
            command=self.sync_upload,
            height=38
        ).pack(side="left", fill="both", expand=True, padx=(0, 6))
        
        # 下载按钮 - 橙色统一风格
        ModernButton(
            quick_sync,
            text="⬇️ 下载",
            style="warning",
            command=self.sync_download,
            height=38
        ).pack(side="left", fill="both", expand=True, padx=(6, 0))
        
        # 功能按钮
        func_btns = ctk.CTkFrame(sync_card, fg_color="transparent")
        func_btns.pack(fill="x", padx=15, pady=(0, 10))
        
        ModernButton(
            func_btns,
            text="📝 修改路径",
            style="primary",
            command=self.edit_paths,
            height=36
        ).pack(side="left", fill="both", expand=True, padx=4)
        
        ModernButton(
            func_btns,
            text="☁️ 云端浏览",
            style="primary",
            command=self.browse_cloud,
            height=36
        ).pack(side="left", fill="both", expand=True, padx=4)
        
        config_btns = ctk.CTkFrame(sync_card, fg_color="transparent")
        config_btns.pack(fill="x", padx=15, pady=(0, 12))
        ModernButton(
            config_btns,
            text="配置保存位置",
            style="secondary",
            command=self.change_config_location,
            height=34
        ).pack(fill="x")

        # 日志卡片
        log_card = ModernCard(self)
        log_card.pack(fill="x", padx=20, pady=(0, 20))
        
        # 新增: 日志标题行
        log_header = ctk.CTkFrame(log_card, fg_color="transparent")
        log_header.pack(fill="x", padx=15, pady=(12, 0))
        
        ctk.CTkLabel(
            log_header,
            text="📋 同步日志",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=Theme.TEXT_GRAY
        ).pack(side="left")
        # 新增: 日志条数
        self.log_count = ctk.CTkLabel(
            log_header,
            text="",
            font=ctk.CTkFont(size=15),
            text_color=Theme.TEXT_MUTED
        )
        self.log_count.pack(side="right")
        
        self.log_text = ctk.CTkTextbox(
            log_card,
            font=ctk.CTkFont(size=15),
            fg_color="#1A1A1A",
            border_color=Theme.BORDER,
            corner_radius=8,
            height=220,
        )
        self.log_text.pack(fill="both", expand=True, padx=15, pady=10)
        self.log_text.bind("<Key>", lambda e: "break")    # 禁止键盘输入
        self.log_text.configure(state="normal")
    
    @staticmethod
    def _normalize_current_index(value, count=None):
        """把保存的 current_account 规范成有效的非负索引。"""
        if count is None:
            count = 0
        try:
            idx = int(value)
        except (TypeError, ValueError):
            idx = 0
        if count <= 0:
            return 0
        return min(max(idx, 0), count - 1)

    def refresh_account_list(self):
        """刷新当前账号显示。"""
        self.current_idx = self._normalize_current_index(self.current_idx, len(self.accounts))
        if self.accounts:
            acc = self.accounts[self.current_idx]
            name = (acc.get('name') or f'账号{self.current_idx + 1}').strip()
            email = (acc.get('user_email') or '未设置').strip()
            local_folder = (acc.get('local_folder') or '').strip()

            self.account_label.configure(text=f"✓ {email}" if email else f"✓ {name}")
            self.folder_label.configure(text=f"📁 {local_folder}" if local_folder else "📁 未设置本地同步文件夹")
            self.account_badge.configure(text=f"账号 {self.current_idx + 1} · {name}")
        else:
            self.account_label.configure(text="未登录")
            self.folder_label.configure(text="")
            self.account_badge.configure(text="")

    def get_current_account(self):
        """获取当前账号，严格限制索引范围。"""
        if not self.accounts:
            self.show_error("请先添加账号")
            return None
        if not 0 <= self.current_idx < len(self.accounts):
            self.current_idx = self._normalize_current_index(self.current_idx, len(self.accounts))
            self.config['current_account'] = self.current_idx
            save_config(self.config)
        return self.accounts[self.current_idx]

    def show_error(self, message: str):
        """显示错误"""
        self.status_indicator.configure(text="●", text_color=Theme.ERROR)
        self.log(f"❌ {message}")

    def show_success(self, message: str):
        """显示成功"""
        self.status_indicator.configure(text="●", text_color=Theme.SUCCESS)
        self.log(f"✅ {message}")

    def show_syncing(self):
        """显示同步中"""
        self.is_syncing = True
        self.status_indicator.configure(text="◐", text_color=Theme.WARNING)

    def queue_log(self, message: str):
        """线程安全地将日志送回 GUI 主线程。"""
        self.log_queue.put(message)

    def process_log_queue(self):
        try:
            while True:
                self.log(self.log_queue.get_nowait())
        except queue.Empty:
            pass
        self.after(100, self.process_log_queue)

    def log(self, message: str):
        """添加日志 - 优化: 带时间戳"""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")

        lines = int(self.log_text.index("end-1c").split(".")[0]) - 1
        self.log_count.configure(text=f"{lines} 条")

    def close_cloud_dialog(self):
        """切换/删除账号前关闭仍绑定旧账号客户端的云端浏览窗口。"""
        dialog = getattr(self, 'cloud_dialog', None)
        if dialog is not None:
            try:
                if dialog.winfo_exists():
                    dialog.destroy()
            except Exception:
                pass
            finally:
                self.cloud_dialog = None

    def switch_account(self):
        """打开账号选择窗口；确认后再切换当前账号。"""
        if self.is_syncing:
            self.show_error("同步进行中，暂不能切换账号")
            return
        if not self.accounts:
            self.show_error("没有账号")
            return
        if len(self.accounts) == 1:
            self.show_success("当前只有一个账号，无需切换")
            return

        existing = getattr(self, "account_dialog", None)
        if existing is not None:
            try:
                if existing.winfo_exists():
                    existing.lift()
                    existing.focus_force()
                    return
            except Exception:
                self.account_dialog = None

        self.current_idx = self._normalize_current_index(self.current_idx, len(self.accounts))
        try:
            self.account_dialog = AccountSelectDialog(
                self,
                self.accounts,
                mode="select",
                current_idx=self.current_idx,
                on_result=self._finish_account_switch,
            )
        except Exception as exc:
            self.account_dialog = None
            logger.exception("打开切换账号窗口失败")
            self.show_error(f"无法打开账号列表: {exc}")

    def _finish_account_switch(self, selected_idx):
        """账号选择窗口关闭后的切换处理。"""
        self.account_dialog = None
        if selected_idx is None:
            return
        if not 0 <= selected_idx < len(self.accounts):
            self.show_error("选择的账号已经不存在")
            return
        if selected_idx == self.current_idx:
            self.show_success("当前账号未改变")
            return

        old_idx = self.current_idx
        self.close_cloud_dialog()
        self.current_idx = selected_idx
        self.config["accounts"] = self.accounts
        self.config["current_account"] = self.current_idx
        try:
            save_config(self.config)
        except Exception as exc:
            self.current_idx = old_idx
            self.config["current_account"] = old_idx
            self.refresh_account_list()
            logger.exception("切换账号保存失败")
            self.show_error(f"切换账号失败，配置保存失败: {exc}")
            return

        self.refresh_account_list()
        self.status_indicator.configure(text="●", text_color="#BBBBBB")
        account = self.accounts[self.current_idx]
        name = str(account.get("name") or f"账号{self.current_idx + 1}")
        email = str(account.get("user_email") or "未获取邮箱")
        folder = str(account.get("local_folder") or "未设置本地文件夹")
        self.show_success(f"已切换到: {name} ({email})")
        self.log(f"   本地文件夹: {folder}")

    def add_account(self):
        """添加账号，并自动让新账号成为当前账号。"""
        if self.is_syncing:
            self.show_error("同步进行中，暂不能添加账号")
            return

        defaults = None
        if self.accounts and 0 <= self.current_idx < len(self.accounts):
            defaults = self.accounts[self.current_idx]
        elif self.accounts:
            defaults = self.accounts[0]

        dialog = AddAccountDialog(self, defaults=defaults)
        self.wait_window(dialog)

        if not dialog.result:
            return

        old_idx = self.current_idx
        self.accounts.append(dialog.result)
        self.current_idx = len(self.accounts) - 1
        self.config['accounts'] = self.accounts
        self.config['current_account'] = self.current_idx
        try:
            save_config(self.config)
        except Exception as exc:
            self.accounts.pop()
            self.current_idx = self._normalize_current_index(old_idx if 'old_idx' in locals() else 0, len(self.accounts))
            self.config['accounts'] = self.accounts
            self.config['current_account'] = self.current_idx
            self.refresh_account_list()
            self.show_error(f"账号添加失败，配置保存失败: {exc}")
            return

        self.close_cloud_dialog()
        self.refresh_account_list()
        new_account = self.accounts[self.current_idx]
        name = new_account.get('name', f'账号{self.current_idx + 1}')
        email = new_account.get('user_email', '未绑定邮箱')
        self.show_success(f"账号添加成功，已切换到: {name} ({email})")

    def delete_account(self):
        """打开删除账号窗口；确认后删除账号。"""
        if self.is_syncing:
            self.show_error("同步进行中，暂不能删除账号")
            return
        if not self.accounts:
            self.show_error("没有账号")
            return

        existing = getattr(self, "account_dialog", None)
        if existing is not None:
            try:
                if existing.winfo_exists():
                    existing.lift()
                    existing.focus_force()
                    return
            except Exception:
                self.account_dialog = None

        self.current_idx = self._normalize_current_index(self.current_idx, len(self.accounts))
        try:
            self.account_dialog = AccountSelectDialog(
                self,
                self.accounts,
                mode="delete",
                current_idx=self.current_idx,
                on_result=self._finish_account_delete,
            )
        except Exception as exc:
            self.account_dialog = None
            logger.exception("打开删除账号窗口失败")
            self.show_error(f"无法打开账号列表: {exc}")

    def _finish_account_delete(self, delete_idx):
        """处理删除账号结果并维护 current_account。"""
        self.account_dialog = None
        if delete_idx is None:
            return
        if not 0 <= delete_idx < len(self.accounts):
            self.show_error("选择的账号已经不存在")
            return

        account_to_delete = self.accounts[delete_idx]
        name = str(account_to_delete.get("name") or f"账号{delete_idx + 1}")
        state_identity = "|".join([
            str(account_to_delete.get("user_email") or ""),
            str(account_to_delete.get("client_id") or ""),
            str(account_to_delete.get("local_folder") or ""),
            str(account_to_delete.get("root_folder") or "/"),
        ])
        state_name = hashlib.sha256(state_identity.encode("utf-8")).hexdigest()[:20] + ".json"
        state_path = os.path.join(APP_DATA, "sync_state", state_name)

        old_idx = self.current_idx
        self.close_cloud_dialog()
        del self.accounts[delete_idx]

        if not self.accounts:
            self.current_idx = 0
        elif delete_idx < old_idx:
            self.current_idx = old_idx - 1
        elif delete_idx == old_idx:
            self.current_idx = min(old_idx, len(self.accounts) - 1)
        else:
            self.current_idx = old_idx

        self.current_idx = self._normalize_current_index(self.current_idx, len(self.accounts))
        self.config["accounts"] = self.accounts
        self.config["current_account"] = self.current_idx

        try:
            save_config(self.config)
        except Exception as exc:
            # 尽量恢复内存中的账号列表，避免磁盘失败后 UI 数据丢失。
            self.accounts.insert(delete_idx, account_to_delete)
            self.current_idx = self._normalize_current_index(old_idx, len(self.accounts))
            self.config["accounts"] = self.accounts
            self.config["current_account"] = self.current_idx
            self.refresh_account_list()
            logger.exception("删除账号保存失败")
            self.show_error(f"删除账号失败，配置保存失败: {exc}")
            return

        try:
            if os.path.exists(state_path):
                os.remove(state_path)
        except OSError as exc:
            logger.warning(f"删除账号同步状态失败: {exc}")

        self.refresh_account_list()
        if self.accounts:
            current = self.accounts[self.current_idx]
            current_name = str(current.get("name") or f"账号{self.current_idx + 1}")
            self.show_success(f"已删除账号: {name}，当前账号: {current_name}")
        else:
            self.show_success(f"已删除账号: {name}，当前已无账号")

    def change_config_location(self):
        """选择配置文件保存目录，并迁移当前配置。"""
        if self.is_syncing:
            self.show_error("同步进行中，暂不能修改配置位置")
            return
        folder = filedialog.askdirectory(title="选择配置文件保存位置", initialdir=APP_DATA)
        if not folder:
            return
        new_dir = os.path.abspath(folder)
        old_dir = os.path.abspath(APP_DATA)
        if new_dir == old_dir:
            self.show_success("配置保存位置未改变")
            return
        try:
            os.makedirs(new_dir, exist_ok=True)
            new_config = os.path.join(new_dir, "config.json")
            old_state_dir = os.path.join(old_dir, "sync_state")
            new_state_dir = os.path.join(new_dir, "sync_state")

            if os.path.exists(new_config):
                if not messagebox.askyesno("目标已有配置", "目标目录已经存在 config.json，是否覆盖？\n选择“否”将取消本次迁移。"):
                    return
            if os.path.isdir(old_state_dir) and os.path.isdir(new_state_dir) and os.listdir(new_state_dir):
                if not messagebox.askyesno("目标已有同步状态", "目标目录已有同步状态。是否删除目标状态后迁移当前状态？"):
                    return

            tmp_config = new_config + ".tmp"
            with open(tmp_config, 'w', encoding='utf-8') as f:
                safe_config = json.loads(json.dumps(self.config, ensure_ascii=False))
                for account in safe_config.get('accounts', []):
                    for key in ('client_secret', 'refresh_token'):
                        if key in account:
                            account[key] = _protect_secret(account[key])
                json.dump(safe_config, f, indent=2, ensure_ascii=False)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp_config, new_config)

            if os.path.isdir(old_state_dir):
                if os.path.isdir(new_state_dir):
                    shutil.rmtree(new_state_dir)
                shutil.copytree(old_state_dir, new_state_dir)

            set_config_dir(new_dir)
            save_config(self.config)
            self.show_success(f"配置保存位置已切换: {new_dir}")
            messagebox.showinfo("配置位置已更新", "配置、同步状态和日志已切换到新位置。旧目录仍保留作为备份。")
        except Exception as e:
            logger.exception("修改配置位置失败")
            self.show_error(f"修改配置位置失败: {e}")

    def edit_paths(self):
        """修改同步路径"""
        account = self.get_current_account()
        if not account:
            return
        
        dialog = EditPathsDialog(self, account)
        self.wait_window(dialog)
        
        if dialog.result:
            self.accounts[self.current_idx] = dialog.result
            self.config['accounts'] = self.accounts
            save_config(self.config)
            self.refresh_account_list()
            self.show_success("路径已更新")
    
    def browse_cloud(self):
        """浏览云端文件"""
        account = self.get_current_account()
        if not account:
            return
        
        client = OneDriveClient(account)
        if not client.load_token():
            self.show_error("Token 无效，请重新授权")
            dialog = AuthDialog(self, account)
            self.wait_window(dialog)
            if not dialog.success:
                return
        
        self.cloud_dialog = CloudBrowserDialog(self, client)
    
    def run_sync(self, mode: str):
        """执行同步"""
        account = self.get_current_account()
        if not account:
            return
        
        if self.is_syncing:
            self.show_error("同步正在进行中")
            return
        
        client = OneDriveClient(account)
        if not client.load_token():
            self.show_error("Token 无效，请重新授权")
            dialog = AuthDialog(self, account)
            self.wait_window(dialog)
            if not dialog.success:
                return
        
        # 保存更新后的 refresh_token
        self.accounts[self.current_idx]['refresh_token'] = client.config['refresh_token']
        save_config(self.config)
        
        # 在新线程中执行同步
        self.stop_event.clear()
        self.show_syncing()
        sync_engine = OneDriveSync(client, account, progress_callback=self.queue_log, stop_event=self.stop_event)
        
        def sync_worker():
            try:
                success = False
                if mode == "full":
                    success = sync_engine.sync_full()
                elif mode == "upload":
                    success = sync_engine.sync_upload()
                elif mode == "download":
                    success = sync_engine.sync_download()
                else:
                    raise ValueError(f"未知同步模式: {mode}")
                if self.stop_event.is_set():
                    self.after(0, self.show_error, "同步已取消")
                elif success:
                    self.after(0, self.show_success, "同步完成")
                else:
                    self.after(0, self.show_error, "同步失败，请查看日志")
            except Exception as e:
                logger.error(f"同步错误: {e}")
                self.after(0, self.show_error, f"同步失败: {str(e)}")
            finally:
                self.after(0, lambda: setattr(self, 'is_syncing', False))
        
        self.sync_thread = threading.Thread(target=sync_worker, daemon=True)
        self.sync_thread.start()
    
    def sync_full(self):
        """完整同步"""
        self.run_sync("full")
    
    def sync_upload(self):
        """仅上传"""
        self.run_sync("upload")
    
    def sync_download(self):
        """仅下载"""
        self.run_sync("download")


# ============== 主程序 ==============

if __name__ == "__main__":
    # 检查依赖
    try:
        import customtkinter
        import requests
    except ImportError as e:
        print(f"缺少依赖包: {e}")
        print("请安装: pip install customtkinter requests")
        sys.exit(1)
    
    # 启动GUI
    app = OneDriveGUI()
    app.mainloop()
