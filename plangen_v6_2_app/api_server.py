from __future__ import annotations

import re
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, JSONResponse

from planner.core.markdown_mapper import parse_markdown_to_template_data
from planner.core.renderer import generate_docx_from_template, enrich_template_context
from planner.core.docx_reader import docx_to_markdown_text
from planner.core.dingtalk_agent_client import call_dingtalk_agent

BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
TEMPLATE_PATH = BASE_DIR / "templates" / "AI_Center_QC_Plan_Template_20260310.docx"

app = FastAPI(title="PlanGen Enterprise API")


@app.get("/")
def root():
    return {"status": "ok", "service": "PlanGen Enterprise API"}


@app.post("/api/generate-plan")
async def generate_plan(
    markdown: str = Form(default=""),
    project_name: str = Form(default=""),
    uploaded_file: UploadFile | None = File(default=None),
):
    try:
        source_text = markdown or ""

        if uploaded_file:
            temp_path = OUTPUT_DIR / uploaded_file.filename
            temp_path.write_bytes(await uploaded_file.read())
            source_text = docx_to_markdown_text(temp_path)

        if not source_text.strip():
            return JSONResponse(status_code=400, content={"success": False, "message": "未检测到输入内容"})

        data = parse_markdown_to_template_data(source_text)
        data = enrich_template_context(data)

        if project_name:
            data["PROJECT_TITLE"] = project_name
            data["project"]["name"] = project_name

        final_name = data.get("PROJECT_TITLE", "稽查计划")
        safe_name = re.sub(r"[\\/:*?\"<>|\r\n]+", "_", final_name).strip()[:60]

        output_path = OUTPUT_DIR / f"{safe_name} 稽查计划.docx"

        generate_docx_from_template(TEMPLATE_PATH, data, output_path)

        return {
            "success": True,
            "project_name": final_name,
            "file_name": output_path.name,
            "download_url": f"/download/{output_path.name}",
        }

    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "message": str(e)})


@app.post("/api/generate-from-agent")
async def generate_from_agent(
    user_prompt: str = Form(...),
    protocol_text: str = Form(default=""),
    project_name: str = Form(default=""),
    uploaded_file: UploadFile | None = File(default=None),
):
    try:
        if uploaded_file:
            temp_path = OUTPUT_DIR / uploaded_file.filename
            temp_path.write_bytes(await uploaded_file.read())
            protocol_text = docx_to_markdown_text(temp_path)

        markdown_result = call_dingtalk_agent(
            prompt=user_prompt,
            protocol_text=protocol_text,
        )

        data = parse_markdown_to_template_data(markdown_result)
        data = enrich_template_context(data)

        if project_name:
            data["PROJECT_TITLE"] = project_name
            data["project"]["name"] = project_name

        final_name = data.get("PROJECT_TITLE", "稽查计划")
        safe_name = re.sub(r"[\\/:*?\"<>|\r\n]+", "_", final_name).strip()[:60]

        output_path = OUTPUT_DIR / f"{safe_name} 稽查计划.docx"

        generate_docx_from_template(TEMPLATE_PATH, data, output_path)

        return {
            "success": True,
            "project_name": final_name,
            "agent_markdown": markdown_result,
            "file_name": output_path.name,
            "download_url": f"/download/{output_path.name}",
        }

    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "message": str(e)})


@app.get("/download/{file_name}")
def download(file_name: str):
    path = OUTPUT_DIR / file_name
    if not path.exists():
        return JSONResponse(status_code=404, content={"success": False, "message": "文件不存在"})
    return FileResponse(path, filename=file_name)
