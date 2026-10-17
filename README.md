# Trợ lý Công tác Xã hội (CTXH Assistant) — Đề tài T5 (HCMUTE)

Hệ thống Trợ lý AI hỗ trợ giải đáp quy định Công tác xã hội (CTXH), tra cứu mã hoạt động 6 chữ số, quy đổi điểm thành ngày CTXH, thẩm định tiến độ tốt nghiệp và đăng ký hoạt động có xác thực thông qua Model Context Protocol (MCP).

---

## 👥 Thành viên nhóm (Nhóm 02FIE)
- **Đoàn Trọng Trung** — MSSV: 24110140 (git: `Dremn`)
- **Khổng Đình Tú** — MSSV: 24110145 (git: `kdteus`)
- **Phạm Thế Việt** — MSSV: 24110147 (git: `vietpham0401`)

---

## 🚀 Hướng dẫn chạy nhanh (1 lệnh)

### 1. Cài đặt môi trường
```bash
make setup
# Hoặc: uv pip install -r requirements.txt
```

### 2. Chạy Demo Terminal Assistant
```bash
make demo
# Hoặc: uv run python rag_client.py
```

**Các lệnh tiện ích trong terminal:**
- `/docs` : Xem danh mục văn bản quy chế trong kho tri thức.
- `/sync` : Quét và đồng bộ tài liệu, sinh viên và hoạt động.
- `/eval` : Đo độ trung thực (Faithfulness) của câu trả lời vừa xong.
- `/clear`: Làm mới lịch sử hội thoại.
- `exit`  : Thoát chương trình.

### 3. Chạy Benchmark Đánh giá (15 test cases)
```bash
make eval
# Hoặc: uv run python tests/evaluate_testset.py
```
Lệnh tự động chạy 15 câu kiểm thử (gồm 4 câu `no_answer` đo Calibration), tính toán Faithfulness, chi phí Token, xuất bảng `eval_report.csv`, báo cáo `eval_report.json` và tự động vẽ 5 biểu đồ chất lượng cao (300 DPI) trong `logs/eval/charts/`.

---

## 🏛️ Kiến trúc 3 tầng bắt buộc

1. **Tầng (a) Tra cứu (Grounding & Calibration):**
   - **Hybrid Search**: Kết hợp Dense Vector (`jina-embeddings-v3`) và từ khóa BM25Okapi.
   - Luôn trích dẫn nguồn văn bản rõ ràng: `[Nguồn: CTXH.md | Chunk ID: ...]`.
   - **Calibration**: Từ chối chính xác các câu hỏi ngoài phạm vi nghiệp vụ (`is_no_answer = true`).

2. **Tầng (b) Biến đổi (Synthesis & Calculation):**
   - **Tra cứu 55 mã hoạt động**: Bóc tách chuẩn hóa mã 6 chữ số (001101, 001202,...).
   - **Quy đổi ngày CTXH**: Áp dụng công thức chia 10 và làm tròn theo quy định trường.
   - **So sánh phiên bản**: Đối chiếu điều kiện tốt nghiệp giữa khóa 2012 (3 ngày) và từ 2013 trở đi (4 ngày).

3. **Tầng (c) Hành động có kiểm chứng (Write + Verify):**
   - Cặp công cụ bắt buộc: `register_ctxh_activity` (Write Tool — cấp mã `REG-CTXH-...` và ghi nhận vào đĩa) ➔ gọi ngay `get_ctxh_status` (Verify Read Tool — đọc lại trạng thái thực tế từ file JSON).

---

## 📂 Cấu trúc thư mục

```
CTXH_Assistant_Rag_Project/
├── data/                       # CSDL văn bản quy định, 55 mã hoạt động, hồ sơ sinh viên
│   ├── CTXH.md                 # Quy định CTXH toàn văn (HCMUTE)
│   ├── activities.json         # Danh mục 55 hoạt động CTXH 6 chữ số
│   ├── students.json           # Danh sách sinh viên
│   └── student_ctxh.json       # Hồ sơ tích lũy CTXH của sinh viên
├── logs/                       # Quản lý tập trung logs và báo cáo
│   ├── app/                    # Logs phiên tương tác thực tế (run_*.jsonl, sessions/)
│   └── eval/                   # Kết quả benchmark (eval_report.json, CSV, charts/)
├── prompts/
│   └── system_agent.txt        # System Prompt ReAct Agent tách biệt khỏi code
├── src/
│   ├── chunking/               # Recursive Character Chunker
│   ├── evaluation/             # Faithfulness Evaluator, Reporter & SessionTracker
│   ├── loaders/                # Document Loaders đa định dạng (docx, pdf, txt, excel) & Metadata
│   ├── retrieval/              # VectorStore (Hybrid BM25 + Jina v3)
│   └── models.py               # Dataclass schemas chặt chẽ
├── tests/
│   ├── testset.jsonl           # Bộ 15 câu hỏi kiểm thử chuẩn hóa
│   └── evaluate_testset.py     # Script benchmark tự động
├── Makefile                    # Lệnh chạy tiện ích (setup, demo, eval)
├── rag_client.py               # Custom ReAct Terminal Client (MAX_STEPS = 6)
├── rag_server.py               # MCP Knowledge & Action Server (giao tiếp stdio)
├── ai-usage.md                 # Báo cáo chi tiết sử dụng trợ lý AI
└── requirements.txt            # Danh sách thư viện phụ thuộc
```
