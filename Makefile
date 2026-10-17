# Makefile for Student Academic Regulations & Procedures Assistant (Final Project)
# Supports standard 1-command deployment: make setup, make demo, make eval

.PHONY: setup demo eval clean

setup:
	@echo [SETUP] 1. Creating local virtual environment (.venv)...
	uv venv
	@echo [SETUP] 2. Installing all dependencies from requirements.txt...
	uv pip install -r requirements.txt
	@echo [SETUP] 3. Initializing .env configuration from .env.example (if not exists)...
	@uv run python -c "import pathlib, shutil; shutil.copy('.env.example', '.env') if not pathlib.Path('.env').exists() and pathlib.Path('.env.example').exists() else None"
	@echo [SETUP] Hoan tat! Hay dam bao GROQ_API_KEY da duoc dien trong file .env truoc khi chay make demo hoac make eval.

demo:
	@echo [DEMO] Launching Terminal Assistant with ReAct Agent Loop...
	uv run python rag_client.py

eval:
	@echo [EVAL] Running 15-testcase benchmark, CSV log and charts generation...
	uv run python tests/evaluate_testset.py

clean:
	uv run python -c "import pathlib; p = pathlib.Path('data/index.json'); p.unlink(missing_ok=True)"

