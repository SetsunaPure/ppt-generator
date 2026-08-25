"""
PPT生成器 - 业务服务层

包含:
- 备课PPT服务 (create_lesson_ppt)
- 讲题PPT服务 (create_topic_ppt)
"""

import json
import logging
from typing import Optional
from datetime import datetime

from .models import LessonPptRequest, TopicPptRequest, EnvConfig, get_env_config
from .outline_parser import OutlineParser, TopicParser, LESSON_OUTLINE_MAP, TOPIC_TEMPLATES
from .builder import PptBuilder
from .utils import (
    get_logger,
    OssOperationHandler,
    send_wechat_notify,
    cleanup_temp_files,
    get_timestamp_filename,
    ensure_dir,
)

logger = get_logger()


# ============================================================
# 备课PPT服务
# ============================================================

def _adapt_lesson_detail(raw_data: dict) -> dict:
    """
    适配API原始格式到OutlineParser期望的格式：
    1. 解包 {code, msg, data} → data
    2. 'outline' → 'outlineList'
    3. 'child' → 'children'（递归）
    """
    # 解包API响应
    if 'data' in raw_data and isinstance(raw_data.get('data'), dict):
        data = raw_data['data']
    else:
        data = raw_data

    def _adapt_node(obj):
        if isinstance(obj, dict):
            result = {}
            for k, v in obj.items():
                if k == 'child':
                    result['children'] = _adapt_node(v)
                else:
                    result[k] = _adapt_node(v)
            return result
        elif isinstance(obj, list):
            return [_adapt_node(item) for item in obj]
        return obj

    data = _adapt_node(data)

    # 映射 outline → outlineList
    if 'outline' in data and 'outlineList' not in data:
        data['outlineList'] = data['outline']

    return data


def create_lesson_ppt(request: LessonPptRequest) -> Optional[str]:
    """
    创建备课PPT

    返回OSS文件链接，异常返回None

    流程：
    1. env = get_env_config(request.activeProfile)
    2. json_data = json.loads(request.lessonDetail)
    3. json_data = _adapt_lesson_detail(json_data)  # 适配API原始格式
    4. pages = OutlineParser().parse(json_data)
    5. ppt_path = PptBuilder(env, pages, request.fontSize,
                             request.fileContentStyle, request.schoolLogo,
                             "lesson").build()
    6. oss_url = OssOperationHandler(oss_region).upload_file(ppt_path, oss_prefix)
    7. 清理临时文件(ppt_path, layout_template_path)
    8. 返回oss_url

    异常处理：
    - logger.error() 记录完整异常
    - send_wechat_notify(env, message=f"lessonId:{request.lessonId} 备课转换异常：{e}")
    - 返回None
    """
    env = None
    ppt_path = None
    layout_path = None

    try:
        # 1. 获取环境配置
        env = get_env_config(request.activeProfile)
        logger.info(f"开始生成备课PPT, lessonId={request.lessonId}, profile={request.activeProfile}")

        # 2. 解析课时大纲
        json_data = json.loads(request.lessonDetail)
        json_data = _adapt_lesson_detail(json_data)  # 适配API原始格式
        logger.debug(f"课时大纲解析完成: title={json_data.get('title')}")

        # 3. 大纲解析
        parser = OutlineParser()
        pages = parser.parse(json_data, font_size=request.fontSize)
        logger.info(f"大纲解析完成，共 {len(pages)} 页")

        # 4. 构建PPT
        builder = PptBuilder(
            env=env,
            pages=pages,
            font_size=request.fontSize,
            file_content_style=request.fileContentStyle,
            logo_path=request.schoolLogo,
            template_prefix="lesson"
        )
        ppt_path = builder.build()
        logger.info(f"PPT构建完成: {ppt_path}")

        # 5. 上传到OSS
        timestamp = get_timestamp_filename("", "")[:13]  # 取时间戳部分
        sub_title = json_data.get("subTitle", "未知")
        oss_prefix = f"lesson/download/pptx/{sub_title}_{timestamp}_{request.fileContentStyle}.pptx"
        oss_handler = OssOperationHandler(env.oss_region)
        oss_url = oss_handler.upload_file(ppt_path, oss_prefix)

        if oss_url:
            logger.info(f"OSS上传成功: {oss_url}")
            return oss_url
        else:
            logger.error("OSS上传失败，返回None")
            return None

    except json.JSONDecodeError as e:
        logger.error(f"课时大纲JSON解析失败: {e}")
        if env:
            send_wechat_notify(env, f"lessonId:{request.lessonId} 备课转换异常：JSON解析失败")
        return None

    except Exception as e:
        logger.error(f"备课PPT生成异常: {e}", exc_info=True)
        if env:
            send_wechat_notify(env, f"lessonId:{request.lessonId} 备课转换异常：{e}")
        return None

    finally:
        # 6. 清理临时文件
        if ppt_path:
            cleanup_temp_files(ppt_path)
        if layout_path:
            cleanup_temp_files(layout_path)


# ============================================================
# 讲题PPT服务
# ============================================================

def create_topic_ppt(request: TopicPptRequest) -> Optional[str]:
    """
    创建讲题PPT

    返回OSS文件链接，异常返回None

    流程：
    1. env = get_env_config(request.activeProfile)
    2. json_data = json.loads(request.detail)
    3. pages = TopicParser(config=TOPIC_TEMPLATES).parse(json_data)
       - 先生成题目页+答案页
       - 再按topicTemplateType匹配模板生成各版块页
    4. ppt_path = PptBuilder(env, pages, request.fontSize,
                             request.fileContentStyle, "",
                             "topic").build()
    5. oss_prefix = f"topic/download/pptx/{request.questionId}_{timestamp}_{request.fileContentStyle}.pptx"
    6. oss_url = OssOperationHandler(env.oss_region).upload_file(ppt_path, oss_prefix)
    7. 清理临时文件
    8. 返回oss_url

    异常处理：同上
    """
    env = None
    ppt_path = None

    try:
        # 1. 获取环境配置
        env = get_env_config(request.activeProfile)
        logger.info(f"开始生成讲题PPT, questionId={request.questionId}, profile={request.activeProfile}")

        # 2. 解析题目详情
        json_data = json.loads(request.detail)
        logger.debug(f"题目详情解析完成: questionId={request.questionId}")

        # 3. 讲题模板解析
        parser = TopicParser(config=TOPIC_TEMPLATES)
        pages = parser.parse(json_data, font_size=request.fontSize, is_origin=False)
        logger.info(f"讲题模板解析完成，共 {len(pages)} 页")

        # 4. 构建PPT
        builder = PptBuilder(
            env=env,
            pages=pages,
            font_size=request.fontSize,
            file_content_style=request.fileContentStyle,
            logo_path="",  # 讲题PPT不使用logo
            template_prefix="topic"
        )
        ppt_path = builder.build()
        logger.info(f"PPT构建完成: {ppt_path}")

        # 5. 上传到OSS
        timestamp = get_timestamp_filename("", "")[:13]
        oss_prefix = f"topic/download/pptx/{request.questionId}_{timestamp}_{request.fileContentStyle}.pptx"
        oss_handler = OssOperationHandler(env.oss_region)
        oss_url = oss_handler.upload_file(ppt_path, oss_prefix)

        if oss_url:
            logger.info(f"OSS上传成功: {oss_url}")
            return oss_url
        else:
            logger.error("OSS上传失败，返回None")
            return None

    except json.JSONDecodeError as e:
        logger.error(f"题目详情JSON解析失败: {e}")
        if env:
            send_wechat_notify(env, f"questionId:{request.questionId} 讲题转换异常：JSON解析失败")
        return None

    except Exception as e:
        logger.error(f"讲题PPT生成异常: {e}", exc_info=True)
        if env:
            send_wechat_notify(env, f"questionId:{request.questionId} 讲题转换异常：{e}")
        return None

    finally:
        # 6. 清理临时文件
        if ppt_path:
            cleanup_temp_files(ppt_path)


# ============================================================
# 状态检查（用于调试/监控）
# ============================================================

def health_check() -> dict:
    """
    健康检查

    返回系统状态信息
    """
    import os
    import sys

    return {
        "status": "healthy",
        "python_version": sys.version,
        "cwd": os.getcwd(),
        "timestamp": datetime.now().isoformat(),
    }
