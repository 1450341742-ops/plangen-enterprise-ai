# PlanGen v6.2 企业级 AI 版

## 免费中国可用 AI 方案
本版默认使用：
- Ollama
- qwen2.5:7b

### 安装
1. 安装 Ollama：https://ollama.com/download
2. 拉取模型：
```powershell
ollama pull qwen2.5:7b
```
3. 启动服务：
```powershell
ollama serve
```

## 启动
```powershell
cd plangen_v6_2_app
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m streamlit run app.py
```

## V6.2 功能
- 基础字段锁死提取
- AI 提取入选标准 / 排除标准 / 流程要求
- AI 一次性返回重点关注，避免逐条调用卡死
- AI 生成 RBQM 策略
- AI 生成访谈问题
- AI 生成发现 / CAPA 草案
- 生成前字段确认页
- Word / PDF 导出

## 建议
如果本机性能一般，可将 `planner/core/ai_engine.py` 中：
```python
OLLAMA_MODEL = "qwen2.5:7b"
```
改成更轻量模型，例如：
```python
OLLAMA_MODEL = "qwen2.5:3b"
```


## V6.2 小模型稳定版修复
- 默认模型改为 `qwen2.5:3b`
- 缩短 AI 输入长度，降低 500 错误概率
- 缩短超时，避免长时间卡住
- 更适合普通 Windows 办公电脑本地运行
- 排查命令改为：
  - `ollama pull qwen2.5:3b`
  - `ollama serve`
  - `ollama run qwen2.5:3b`
