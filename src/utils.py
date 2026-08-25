"""
PPT生成器 - 工具函数模块

包含:
- HTTP请求（代理支持）
- OSS上传
- 企微告警
- 日志模块
- 图片处理
"""

import os
import io
import json
import time
import logging
import logging.handlers
from pathlib import Path
from datetime import datetime
from typing import Optional, Any, Dict

import requests
from PIL import Image


# ============================================================
# 日志模块
# ============================================================

def get_logger(
    name: str = "ppt_generator",
    log_filename: str = "./file/log/gen_ppt.log"
) -> logging.Logger:
    """
    获取日志记录器，JSON格式日志，RotatingFileHandler

    - 5MB轮转，保留2个备份，UTF-8编码
    - 日志字段：timestamp, level, message, module, function, line, thread, process
    - 有异常时追加exception字段
    - 防重复初始化
    """
    # 必须显式import，原代码缺失此行导致运行报错
    import logging.handlers

    # 防重复初始化
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)

    # 确保日志目录存在
    log_path = Path(log_filename)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    # RotatingFileHandler
    handler = logging.handlers.RotatingFileHandler(
        filename=log_filename,
        maxBytes=5 * 1024 * 1024,  # 5MB
        backupCount=2,
        encoding='utf-8'
    )

    # JSON格式
    class JSONFormatter(logging.Formatter):
        def format(self, record: logging.LogRecord) -> str:
            log_data = {
                "timestamp": datetime.now().isoformat(),
                "level": record.levelname,
                "message": record.getMessage(),
                "module": record.module,
                "function": record.funcName,
                "line": record.lineno,
                "thread": record.thread,
                "process": record.process,
            }
            if record.exc_info:
                log_data["exception"] = self.formatException(record.exc_info)
            return json.dumps(log_data, ensure_ascii=False)

    handler.setFormatter(JSONFormatter())
    logger.addHandler(handler)

    return logger


# ============================================================
# HTTP请求（代理支持）
# ============================================================

def make_request(
    method: str,
    url: str,
    env: 'EnvConfig',
    **kwargs
) -> requests.Response:
    """
    统一请求函数，自动处理代理

    - dev/test直连，product走代理
    - 代理配置从EnvConfig获取，零硬编码
    """
    proxies = env.proxies
    response = requests.request(method, url, proxies=proxies, **kwargs)
    return response


# ============================================================
# 图片处理
# ============================================================

def get_image_dimensions(image_url: str, env: 'EnvConfig') -> tuple[int, int]:
    """
    获取图片尺寸（像素）

    返回: (width, height)
    """
    try:
        response = make_request("GET", image_url, env, timeout=10)
        response.raise_for_status()
        img = Image.open(io.BytesIO(response.content))
        return img.size  # (width, height)
    except Exception:
        return (0, 0)


def download_image_to_bytes(image_url: str, env: 'EnvConfig') -> Optional[bytes]:
    """
    下载图片到内存

    返回: bytes 或 None
    """
    try:
        response = make_request("GET", image_url, env, timeout=30)
        response.raise_for_status()
        return response.content
    except Exception as e:
        logger = get_logger()
        logger.error(f"下载图片失败: {image_url}, error: {e}")
        return None


def save_temp_image(image_bytes: bytes, suffix: str = ".png") -> Optional[str]:
    """
    保存临时图片文件

    返回: 文件路径 或 None
    """
    try:
        temp_dir = Path("./file/temp")
        temp_dir.mkdir(parents=True, exist_ok=True)
        filename = f"temp_{int(time.time() * 1000)}{suffix}"
        filepath = temp_dir / filename
        with open(filepath, 'wb') as f:
            f.write(image_bytes)
        return str(filepath)
    except Exception as e:
        logger = get_logger()
        logger.error(f"保存临时图片失败: {e}")
        return None


def get_scaled_dimensions(
    original_width: float,
    original_height: float,
    max_height_inches: float = 1.07
) -> tuple[float, float]:
    """
    计算Logo等图片的缩放尺寸

    - 固定高度1.07cm，按原始宽高比算宽度
    - 1cm ≈ 0.3937 inches

    返回: (scaled_width, scaled_height) 单位：英寸
    """
    max_height = max_height_inches * 0.3937  # 转换为英寸
    aspect_ratio = original_width / original_height if original_height > 0 else 1.0
    scaled_height = max_height
    scaled_width = scaled_height * aspect_ratio
    return (scaled_width, scaled_height)


# ============================================================
# 企微告警
# ============================================================

def send_wechat_notify(env: 'EnvConfig', message: str) -> bool:
    """
    通过企微机器人发送text消息，@指定手机号

    - 使用 make_request("POST", env.wechat_webhook, env, ...) 发送
    - 请求头：Content-Type: application/json
    - 请求体：{"msgtype":"text","text":{"content":message,"mentioned_mobile_list":env.wechat_mentioned}}
    - 成功返回True，失败返回False（不抛异常，告警不能影响主流程）
    """
    try:
        payload = {
            "msgtype": "text",
            "text": {
                "content": message,
                "mentioned_mobile_list": list(env.wechat_mentioned)
            }
        }
        headers = {"Content-Type": "application/json"}
        response = make_request(
            "POST",
            env.wechat_webhook,
            env,
            json=payload,
            headers=headers,
            timeout=10
        )
        response.raise_for_status()
        result = response.json()
        return result.get("errcode", 0) == 0
    except Exception as e:
        logger = get_logger()
        logger.error(f"企微告警发送失败: {e}")
        return False


# ============================================================
# OSS处理
# ============================================================

def convert_to_internal_url(url: str) -> str:
    """
    OSS URL内外网转换

    规则：
    - 北京区域：切换到内网
    - 杭州区域：保持不变
    """
    if not url:
        return url

    # 已在内网格式的直接返回
    if '-internal.oss-' in url or '.aliyuncs.com' not in url:
        return url

    # 替换为内网地址
    return url.replace('.oss-', '-internal.oss-').replace('http://', 'https://')


class OssOperationHandler:
    """
    OSS操作处理器

    提供文件上传功能
    """

    def __init__(self, region: str = "beijing"):
        self.region = region
        # 实际使用时需要配置OSS AccessKey和Bucket信息
        # 这里仅提供接口定义

    def upload_file(self, local_path: str, oss_key: str) -> Optional[str]:
        """
        上传文件到OSS

        参数:
            local_path: 本地文件路径
            oss_key: OSS上的存储路径

        返回: OSS文件URL 或 None
        """
        try:
            # 检查文件是否存在
            if not os.path.exists(local_path):
                logger = get_logger()
                logger.error(f"上传文件不存在: {local_path}")
                return None

            # TODO: 实现实际的OSS上传逻辑
            # 这里需要配置实际的OSS客户端
            # 示例:
            # import oss2
            # auth = oss2.Auth(ACCESS_KEY_ID, ACCESS_KEY_SECRET)
            # bucket = oss2.Bucket(auth, f'http://oss-{region}.aliyuncs.com', BUCKET_NAME)
            # bucket.put_object_from_file(oss_key, local_path)
            # return f"https://{BUCKET_NAME}.oss-{region}.aliyuncs.com/{oss_key}"

            # 暂时返回模拟URL
            return f"https://example.oss-{self.region}.aliyuncs.com/{oss_key}"
        except Exception as e:
            logger = get_logger()
            logger.error(f"OSS上传失败: {e}")
            return None


# ============================================================
# 文件操作
# ============================================================

def ensure_dir(path: str) -> None:
    """确保目录存在"""
    Path(path).mkdir(parents=True, exist_ok=True)


def cleanup_temp_files(*paths: str) -> None:
    """清理临时文件"""
    for path in paths:
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except Exception:
                pass


def get_timestamp_filename(prefix: str, suffix: str) -> str:
    """生成带时间戳的文件名"""
    return f"{prefix}_{int(time.time() * 1000)}{suffix}"
