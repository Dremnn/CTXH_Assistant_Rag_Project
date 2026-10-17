# AI usage declaration

**Team Members:**
- **Đoàn Trọng Trung** (Student ID: 24110140, git: `Dremn`)
- **Khổng Đình Tú** (Student ID: 24110145, git: `kdteus`)
- **Phạm Thế Việt** (Student ID: 24110147, git: `vietpham0401`)

**Course:** AI Programming with Python (02FIE) — HCMUTE  
**Project:** T5 — Social Work Regulations & Procedures Assistant (CTXH Assistant)  
**Tools:** Google Antigravity & Claude Code (Anthropic), used in VS Code, October 2026.

---

## How I used AI

- **Choosing the tools and architecture:** AI suggested the core MCP tools needed to satisfy the 3 mandatory architectural layers: grounding and student verification (`search_ctxh_rules`, `check_student_exists`), calculation and synthesis (`get_activity`, `convert_points_to_ctxh_days`, `check_graduation_eligibility`), and the mandatory action pair `register_ctxh_activity` (write tool with transaction ID) followed immediately by `get_ctxh_status` (verification read tool).
- **Document Loaders and Metadata:** AI helped design the loader pipeline in `src/loaders/` supporting `.docx`, `.pdf`, `.txt`, and `.xlsx`. AI provided the regular expressions and parsing logic for extracting YAML front-matter headers (`--- doc_id ... ---`) from documents, as well as filename regex fallbacks (`<doc_id>_<year>.<ext>`) in `metadata.py`.
- **Data curation and modeling:** AI assisted in parsing and structuring `data/activities.json` by extracting 55 standardized 6-digit activity codes from `data/CTXH.md`. It also suggested schemas for the student databases (`students.json` and `student_ctxh.json`) and typed `@dataclass` models in `src/models.py`.
- **Logging management and Telemetry:** AI guided the telemetry architecture in `src/evaluation/tracker.py` (`SessionTracker`), structuring separate JSONL execution traces for demo sessions (`logs/app/run_*.jsonl`) and benchmark runs (`logs/eval/run_eval_*.jsonl`). It recorded `prompt_tokens`, `completion_tokens`, latency, tool call sequences, and exported session summaries to CSV.
- **Chart generation (Visualization):** AI wrote plotting routines using `matplotlib` and `seaborn` in `src/evaluation/reporter.py` to automatically output 5 high-resolution (300 DPI) charts: token usage distribution (`01_token_usage.png`), response latency vs tool calls (`02_latency_and_calls.png`), quality metrics (`03_quality_metrics.png`), overall evaluation dashboard (`04_overall_dashboard.png`), and Dense vs Hybrid search comparison (`05_comparison_dense_vs_hybrid.png`).
- **Testset (15 benchmark questions):** AI supported drafting the 15 benchmark questions in `tests/testset.jsonl`, covering 6 functional categories: logic rules, version comparison (2012 vs 2013+ cohort), point-to-day conversion calculations, action write-and-verify, missing information, and exactly 4 out-of-domain `no_answer` questions for calibration. We cross-checked and verified all expected answers against HCMUTE regulations.
- **Writing and fixing the code:** We wrote the custom ReAct agent loop and MCP server logic ourselves. AI reviewed our code and helped resolve errors, such as Windows UTF-8 terminal encoding errors (`UnicodeEncodeError`), tool call JSON markdown parsing, Groq API rate-limit exponential backoff, and BM25/Dense weight tuning in `vector_store.py`.
- **README and Makefile:** AI helped generate a clean, streamlined `README.md` and `Makefile` so the entire project can be reproduced with standard single commands (`make setup`, `make demo`, `make eval`).

---

## Part / AI role / Written by

| Part | AI role | Written by |
| :--- | :--- | :--- |
| **Tool selection & Architecture** | Suggested MCP tools matching the 3 mandatory layers and ReAct pattern | Decision & design by team, suggested by AI |
| **Action & Verify Tools (`register_ctxh_activity`, `get_ctxh_status`)** | Explained write-then-verify pattern; suggested transaction ID generation format | Code by Trung, review & bugfix by AI |
| **Student Tools (`check_student_exists`, `check_graduation_eligibility`)** | Explained validation rules and graduation threshold comparison logic | Code by Tú, logic verified by AI |
| **Retrieval & Rules Tools (`search_ctxh_rules`, `get_activity`)** | Guided version filtering logic (newest version priority and year parameter) | Code by Việt & Tú |
| **Document Loaders & Metadata (`src/loaders/`)** | Suggested multi-format loaders (docx, pdf, txt, excel) and front-matter regex | Code by Tú, regex suggested by AI |
| **Data curation (`CTXH.md`, `activities.json`, `students.json`)** | Extracted and normalized 55 activity codes into structured JSON | Extracted with AI, verified by team against policy |
| **Chart generation (`reporter.py`)** | Wrote matplotlib/seaborn visualization functions for 5 high-DPI charts | Functions drafted by AI, customized by Việt |
| **Logging & Telemetry (`tracker.py`)** | Designed JSONL schema, SessionTracker, and CSV exporter | Code by Việt & Trung, format guided by AI |
| **Benchmark Testset (`tests/testset.jsonl` - 15 questions)** | Drafted 15 test questions including 4 `no_answer` calibration cases | Drafted with AI, checked and finalized by team |
| **ReAct Agent Loop (`rag_client.py`)** | Provided clean ReAct loop scaffold with Groq LLM integration and error handling | Written by Trung, bugfixes suggested by AI |
| **README.md & Makefile** | Simplified README and configured standard 1-command deployment | Written by AI, checked and tested by team |

---

## Main prompts (examples)

- *"Suggest the required MCP tools for a Social Work Assistant project that needs to satisfy 3 layers: grounding, calculation, and a write action with mandatory verification."*
- *"Help me write a Python metadata extractor in `src/loaders/metadata.py` that reads YAML front-matter (`--- doc_id ... ---`) from docx and md files, with regex fallback based on filename `<doc_id>_<year>.<ext>`."*
- *"Create a matplotlib script in `reporter.py` to generate 4 high-DPI (300 DPI) charts for RAG evaluation: token usage distribution, latency vs tool calls, quality metrics, and a 2x2 dashboard."*
- *"Design a `SessionTracker` class that tracks agent thoughts, tool arguments, output, token usage, and exports both JSONL and CSV logs."*
- *"Help draft 15 evaluation questions in `testset.jsonl` based on HCMUTE social work regulations, with exactly 4 out-of-domain questions (`is_no_answer = true`) to evaluate calibration."*
- *"Why is my Windows terminal throwing `UnicodeEncodeError: 'charmap'` when printing agent responses, and how can I safely reconfigure UTF-8 stdout?"*
- *"Simplify our README.md so that the instructor can clone, setup, and run `make demo` and `make eval` with one command."*
