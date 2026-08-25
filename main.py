"""
PPT生成器 - FastAPI入口

端口8001，root_path=/api
接口:
  POST /lesson/create_ppt    备课PPT生成
  POST /exam/topic/create_ppt 讲题PPT生成
  GET  /health               健康检查
"""

from fastapi import FastAPI, Body, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field
from typing import Optional

from src.service import create_lesson_ppt, create_topic_ppt, health_check
from src.models import ENV_PRESETS


# ============================================================
# Pydantic请求模型（API层校验）
# ============================================================

class LessonPptRequestBody(BaseModel):
    """备课PPT请求体"""
    lessonId: str = Field(..., description="课时ID")
    stage: int = Field(default=12, description="学段")
    subject: int = Field(default=9912, description="学科")
    lessonDetail: str = Field(..., description="课时大纲JSON字符串")
    fontSize: int = Field(default=16, description="字号")
    activeProfile: str = Field(default="dev", description="环境标识: dev/test/product")
    fileContentStyle: str = Field(default="0", description="模板风格")
    schoolLogo: str = Field(default="", description="学校Logo URL")


class TopicPptRequestBody(BaseModel):
    """讲题PPT请求体"""
    questionId: str = Field(..., description="题目ID")
    stage: int = Field(default=12, description="学段")
    subject: int = Field(default=9912, description="学科")
    detail: str = Field(..., description="题目详情JSON字符串")
    fontSize: int = Field(default=16, description="字号")
    activeProfile: str = Field(default="dev", description="环境标识: dev/test/product")
    fileContentStyle: str = Field(default="0", description="模板风格")


# ============================================================
# FastAPI应用
# ============================================================

app = FastAPI(
    title="备课PPT生成服务",
    description="教育备课PPT自动生成微服务，接收课时大纲JSON，自动生成PPT并上传OSS",
    version="1.2.0",
    root_path="/api",
)


@app.post("/lesson/create_ppt", summary="备课PPT生成")
def api_create_lesson_ppt(request: LessonPptRequestBody = Body(...)):
    """备课PPT生成接口"""
    if request.activeProfile not in ENV_PRESETS:
        raise HTTPException(
            status_code=400,
            detail=f"无效的activeProfile: {request.activeProfile}，可选: {list(ENV_PRESETS.keys())}"
        )

    from src.models import LessonPptRequest
    internal_req = LessonPptRequest(
        lessonId=request.lessonId,
        stage=request.stage,
        subject=request.subject,
        lessonDetail=request.lessonDetail,
        fontSize=request.fontSize,
        activeProfile=request.activeProfile,
        fileContentStyle=request.fileContentStyle,
        schoolLogo=request.schoolLogo,
    )

    result = create_lesson_ppt(internal_req)

    if result is None:
        return JSONResponse(
            status_code=500,
            content={"code": 500, "message": "备课PPT生成失败，请查看服务日志", "data": None}
        )

    return {"code": 200, "message": "success", "data": result}


@app.post("/exam/topic/create_ppt", summary="讲题PPT生成")
def api_create_topic_ppt(request: TopicPptRequestBody = Body(...)):
    """讲题PPT生成接口"""
    if request.activeProfile not in ENV_PRESETS:
        raise HTTPException(
            status_code=400,
            detail=f"无效的activeProfile: {request.activeProfile}，可选: {list(ENV_PRESETS.keys())}"
        )

    from src.models import TopicPptRequest
    internal_req = TopicPptRequest(
        questionId=request.questionId,
        stage=request.stage,
        subject=request.subject,
        detail=request.detail,
        fontSize=request.fontSize,
        activeProfile=request.activeProfile,
        fileContentStyle=request.fileContentStyle,
    )

    result = create_topic_ppt(internal_req)

    if result is None:
        return JSONResponse(
            status_code=500,
            content={"code": 500, "message": "讲题PPT生成失败，请查看服务日志", "data": None}
        )

    return {"code": 200, "message": "success", "data": result}


@app.get("/health", summary="健康检查")
def api_health_check():
    """服务健康检查"""
    return health_check()


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8001)
