import os
import sys
import csv
import json
import re
import uuid
from pathlib import Path
from datetime import datetime
from typing import Optional, Any
from dotenv import load_dotenv
from mcp.server.mcpserver import MCPServer

# Set UTF-8 encoding for standard streams
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

# Add project root to sys.path
_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

from openpyxl import load_workbook

from src.loaders import DocumentLoaderFactory, SUPPORTED_EXTENSIONS
from src.chunking import RecursiveChunker
from src.retrieval import VectorStore, EmbeddingGenerator

load_dotenv()

# Initialize MCP Server
mcp = MCPServer("rag-knowledge-server")

DATA_DIR = _HERE / "data"
INDEX_PATH = DATA_DIR / "index.json"

# Ho tro ca activities.json va ctxh.json
ACTIVITIES_RULES_PATHS = [
    DATA_DIR / "activities.json",
]
CTXH_RULES_JSON_PATH = DATA_DIR / "activities.json"
CTXH_RULES_DOC_EXTENSIONS = {".xlsx", ".xlsm", ".docx", ".doc", ".pdf", ".txt", ".md"}

# Cac file Excel luu hoat dong sinh vien da tham gia
CTXH_STUDENT_ACTIVITY_FILENAMES = [
    "student_activities.xlsx",
]

# Global Store
STORE = VectorStore()


def _get_activities_rules_path() -> Path:
    for p in ACTIVITIES_RULES_PATHS:
        if p.exists():
            return p
    return ACTIVITIES_RULES_PATHS[0]


def sync_knowledge_base() -> dict:
    """
    Incremental synchronization of data/ directory:
    - Finds new files -> chunks, embeds, adds to store.
    - Finds modified files -> removes old chunks, re-chunks, re-embeds.
    - Finds deleted files -> removes chunks from store.
    - Saves updated store to index.json.
    """
    global STORE
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    # 1. Discover all supported files currently on disk in data/
    current_files = {}
    for p in DATA_DIR.glob("*"):
        if p.name.startswith("~$") or p.name.startswith("."):
            continue
        ignored_excel = {name.lower() for name in CTXH_STUDENT_ACTIVITY_FILENAMES} | {"students_ctxh.xlsx", "students.xlsx"}
        if p.name.lower() in ignored_excel:
            continue
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS:
            current_files[p.name] = (p, p.stat().st_mtime)

    # Populate file_meta if missing from existing chunks
    if not STORE.file_meta and STORE.chunks:
        for chunk in STORE.chunks.values():
            fname = chunk.metadata.get("file_name")
            if fname and fname in current_files and fname not in STORE.file_meta:
                STORE.file_meta[fname] = current_files[fname][1]

    added_files = []
    updated_files = []
    removed_files = []

    # 2. Check for deleted files (was in store, but no longer on disk)
    for indexed_file in list(STORE.file_meta.keys()):
        if indexed_file not in current_files:
            STORE.remove_chunks_by_file(indexed_file)
            removed_files.append(indexed_file)

    # 3. Check for new or modified files
    emb_gen = None
    chunker = RecursiveChunker(chunk_size=600, chunk_overlap=80)

    for fname, (fpath, mtime) in current_files.items():
        is_new = fname not in STORE.file_meta
        is_modified = not is_new and mtime > STORE.file_meta[fname]

        if is_new or is_modified:
            try:
                docs = DocumentLoaderFactory.load_file(fpath)
                chunks = chunker.split_documents(docs)
            except Exception as e:
                sys.stderr.write(f"[Warning] Failed to load/chunk {fname}: {e}\n")
                continue

            if is_modified:
                STORE.remove_chunks_by_file(fname)
                updated_files.append(fname)
            else:
                added_files.append(fname)

            if chunks:
                embeddings = None
                if emb_gen is None:
                    try:
                        emb_gen = EmbeddingGenerator()
                    except Exception as e:
                        sys.stderr.write(f"[Warning] EmbeddingGenerator unavailable: {e}\n")
                        emb_gen = None

                if emb_gen is not None:
                    try:
                        texts = [c.content for c in chunks]
                        embeddings = emb_gen.embed_documents(texts, batch_size=16)
                    except Exception as e:
                        sys.stderr.write(f"[Warning] Failed to compute embeddings for {fname}: {e}\n")
                        embeddings = None

                STORE.add_chunks(chunks, embeddings)

            STORE.file_meta[fname] = mtime

    # 4. Save to index.json if any changes occurred
    if added_files or updated_files or removed_files or not INDEX_PATH.exists():
        STORE.save_to_json(INDEX_PATH)

    return {
        "added": added_files,
        "updated": updated_files,
        "removed": removed_files,
        "total_chunks": len(STORE.chunks),
        "total_files": len(STORE.file_meta),
    }


def _ensure_index_loaded() -> None:
    """Load cached index, sync students from excel, extract CTXH rules to json, and synchronize data/."""
    global STORE
    if INDEX_PATH.exists():
        try:
            STORE = VectorStore.load_from_json(INDEX_PATH)
        except Exception:
            STORE = VectorStore()

    sync_knowledge_base()
    sync_ctxh_rules_from_file()


def _extract_rules_from_text(text: str, rules_dict: dict) -> None:
    """Boc tach quy tac tinh diem, tru diem va loai tru tu van ban quy che."""
    t_lower = text.lower()
    if "chia cho 10" in t_lower:
        rules_dict["formula"] = "Lấy tổng điểm hoạt động CTXH SV đã tích lũy chia cho 10 (số lẻ dưới 5 làm tròn xuống; tròn 5 điểm, quy đổi thành nửa ngày CTXH; trên 5 làm tròn thành 10, quy đổi thành một ngày CTXH)."
    elif "chia cho 12" in t_lower and "formula" not in rules_dict:
        rules_dict["formula"] = "Lấy tổng điểm CTXH tích lũy chia cho 12 (chuẩn quốc tế áp dụng từ 2025)."
    elif "chia cho 8" in t_lower and "formula" not in rules_dict:
        rules_dict["formula"] = "Lấy tổng điểm hoạt động CTXH tích lũy chia cho 8 (Quy định 2023)."

    for line in text.splitlines():
        l_lower = line.lower()
        if ("trừ" in l_lower or "tru " in l_lower) and ("ngày" in l_lower or "ngay" in l_lower or "đrl" in l_lower or "kỷ luật" in l_lower):
            clean = line.strip().lstrip("-*•✓ ").strip()
            if len(clean) > 15 and clean not in rules_dict.setdefault("deduction_rules", []):
                rules_dict["deduction_rules"].append(clean)
        if ("không được tính" in l_lower or "không tính" in l_lower or "thù lao" in l_lower or "tiền công" in l_lower):
            clean = line.strip().lstrip("-*•✓ ").strip()
            if len(clean) > 15 and clean not in rules_dict.setdefault("exclusions", []):
                rules_dict["exclusions"].append(clean)


def _find_all_ctxh_rules_files() -> list[Path]:
    """
    Tim tat ca cac file quy che va bang diem CTXH (.xlsx, .xlsm, .docx, .doc, .pdf, .txt, .md) trong thu muc data/.
    Bo hoan toan viec boc tach tu file .json ma thay bang Excel, Word, PDF va Text.
    """
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    files: list[Path] = []

    # Danh sach cac file Excel sinh vien / hoat dong sinh vien can loai tru khoi tap van ban quy che
    student_excel_names = {
        name.lower() for name in CTXH_STUDENT_ACTIVITY_FILENAMES
    } | {"students.xlsx", "students_ctxh.xlsx", "danh_sach_sinh_vien.xlsx"}

    # Quet cac file quy che trong data/ (tuyet doi khong doc file .json)
    for p in sorted(DATA_DIR.glob("*")):
        if p.name.startswith("~$") or p.name.startswith("."):
            continue
        # Bo qua tat ca cac file JSON (khong dung json de boc tach nua)
        if p.suffix.lower() == ".json":
            continue
        # Bo qua file Excel danh sach sinh vien va hoat dong da tham gia
        if p.name.lower() in student_excel_names:
            continue
        if p.is_file() and p.suffix.lower() in CTXH_RULES_DOC_EXTENSIONS:
            if p not in files:
                files.append(p)

    return files


def _find_ctxh_rules_source_file() -> Path | None:
    """Tim kiem mot file quy che CTXH uu tien trong data/."""
    files = _find_all_ctxh_rules_files()
    return files[0] if files else None


def _parse_ctxh_rules_from_file(fpath: Path) -> dict:
    """
    Boc tach noi dung ve danh muc hoat dong va cach tinh diem CTXH tu cac file doc, docx, pdf, json, txt.
    Ho tro da dinh dang:
      - .json: doc danh muc hoat dong va calculation_rules truc tiep.
      - .docx / .doc: doc bang bieu va cac doan van quy dinh.
      - .pdf: trich xuat text va cac bang bieu ma 6 chu so qua pypdf.
      - .txt / .md: doc bang markdown (|) va noi dung quy che.
    Tra ve dict gom calculation_rules va danh sach activities.
    """
    ext = fpath.suffix.lower()
    activities: list[dict] = []
    calculation_rules = {
        "formula": "Lấy tổng điểm hoạt động CTXH SV đã tích lũy chia cho 10 (số lẻ dưới 5 làm tròn xuống; tròn 5 điểm, quy đổi thành nửa ngày CTXH; trên 5 làm tròn thành 10, quy đổi thành một ngày CTXH).",
        "deduction_rules": [
            "SV đăng ký tham gia hoạt động CTXH nhưng không thực hiện mà không có lý do hoặc lý do không chính đáng bị trừ 0,5 đến 2,0 ngày CTXH và hạ bậc ĐRL.",
            "SV không hoàn thành nhiệm vụ, gây ảnh hưởng xấu bị trừ từ 01 đến toàn bộ số ngày đã tích lũy.",
            "SV gian lận công nghệ (Fake GPS, can thiệp camera điểm danh) bị hủy bỏ toàn bộ số ngày CTXH đã tích lũy và kỷ luật."
        ],
        "exclusions": [
            "Các hoạt động không mang tính chất CTXH như xem biểu diễn văn nghệ, xem thể thao, du lịch, tham quan, sinh hoạt CLB không được tính ngày CTXH.",
            "Khi SV tham gia hoạt động CTXH đã được trả tiền công hoặc trả phí thì không được tính ngày CTXH."
        ]
    }


    if ext in {".docx", ".doc"}:
        try:
            from docx import Document
            doc = Document(str(fpath))
            curr_cat = ""
            for table in doc.tables:
                for row in table.rows:
                    cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                    for idx, val in enumerate(cells):
                        if re.match(r"^\d{6}$", val) and len(cells) >= idx + 3:
                            code = val
                            cat = cells[0] if idx > 0 and cells[0] else curr_cat
                            if cat and cat != code:
                                curr_cat = cat
                            desc = cells[idx + 1] if len(cells) > idx + 1 else ""
                            points = cells[idx + 2] if len(cells) > idx + 2 else ""
                            days = cells[idx + 3] if len(cells) > idx + 3 else ""
                            activities.append({
                                "activity_code": code,
                                "category": curr_cat,
                                "activity_name": desc,
                                "points": points,
                                "converted_days": days,
                                "formula": "Điểm / 10 = Ngày CTXH"
                            })
                            break
            full_doc_text = "\n".join(p.text for p in doc.paragraphs)
            _extract_rules_from_text(full_doc_text, calculation_rules)
        except Exception as e:
            sys.stderr.write(f"[sync_ctxh_rules] Loi doc docx {fpath.name}: {e}\n")

    elif ext == ".pdf":
        try:
            from pypdf import PdfReader
            reader = PdfReader(str(fpath))
            full_pdf_text = []
            for page in reader.pages:
                txt = page.extract_text() or ""
                full_pdf_text.append(txt)
            all_text = "\n".join(full_pdf_text)

            curr_cat = ""
            for line in all_text.splitlines():
                line_str = line.strip()
                if not line_str:
                    continue
                m = re.search(r"\b(\d{6})\b", line_str)
                if m:
                    code = m.group(1)
                    if "|" in line_str:
                        parts = [c.strip() for c in line_str.split("|")]
                        if parts and parts[0] == "":
                            parts = parts[1:]
                        if parts and parts[-1] == "":
                            parts = parts[:-1]
                        for idx, p in enumerate(parts):
                            if p == code and len(parts) >= idx + 3:
                                cat = parts[0] if idx > 0 and parts[0] else curr_cat
                                if cat and cat != code:
                                    curr_cat = cat
                                desc = parts[idx + 1]
                                points = parts[idx + 2]
                                days = parts[idx + 3] if len(parts) > idx + 3 else ""
                                activities.append({
                                    "activity_code": code,
                                    "category": curr_cat,
                                    "activity_name": desc,
                                    "points": points,
                                    "converted_days": days,
                                    "formula": "Điểm / 10 = Ngày CTXH"
                                })
                                break
                    else:
                        parts = line_str.split(code, 1)
                        prefix = parts[0].strip()
                        suffix = parts[1].strip()
                        if prefix and not any(ch.isdigit() for ch in prefix):
                            curr_cat = prefix
                        score_matches = list(re.finditer(r"(\d+(?:[.,]\d+)?(?:\s*[→\->]\s*\d+(?:[.,]\d+)?)?)", suffix))
                        if len(score_matches) >= 2:
                            points_match = score_matches[-2].group(1).strip()
                            days_match = score_matches[-1].group(1).strip()
                            desc_end = score_matches[-2].start()
                            desc = suffix[:desc_end].strip().strip("-:| ")
                            activities.append({
                                "activity_code": code,
                                "category": curr_cat,
                                "activity_name": desc or f"Hoạt động CTXH {code}",
                                "points": points_match,
                                "converted_days": days_match,
                                "formula": "Điểm / 10 = Ngày CTXH"
                            })
            _extract_rules_from_text(all_text, calculation_rules)
        except Exception as e:
            sys.stderr.write(f"[sync_ctxh_rules] Loi doc pdf {fpath.name}: {e}\n")

    elif ext in {".md", ".txt"}:
        try:
            with open(fpath, "r", encoding="utf-8") as f:
                content = f.read()

            curr_cat = ""
            for line in content.splitlines():
                if "|" not in line:
                    continue
                parts = [c.strip() for c in line.split("|")]
                code_idx = -1
                for idx, p in enumerate(parts):
                    if re.match(r"^\d{6}$", p):
                        code_idx = idx
                        break
                if code_idx != -1:
                    code = parts[code_idx]
                    cat_candidates = [parts[i] for i in range(code_idx) if parts[i]]
                    if cat_candidates:
                        curr_cat = cat_candidates[-1]
                    desc = parts[code_idx + 1] if len(parts) > code_idx + 1 else ""
                    points = parts[code_idx + 2] if len(parts) > code_idx + 2 else ""
                    days = parts[code_idx + 3] if len(parts) > code_idx + 3 else ""
                    activities.append({
                        "activity_code": code,
                        "category": curr_cat,
                        "activity_name": desc,
                        "points": points,
                        "converted_days": days,
                        "formula": "Điểm / 10 = Ngày CTXH"
                    })
            _extract_rules_from_text(content, calculation_rules)
        except Exception as e:
            sys.stderr.write(f"[sync_ctxh_rules] Loi doc md/txt {fpath.name}: {e}\n")

    elif ext in {".xlsx", ".xlsm"}:
        try:
            wb = load_workbook(str(fpath), data_only=True)
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                curr_cat = ""
                for row in ws.iter_rows(values_only=True):
                    row_strs = [str(c).strip() if c is not None else "" for c in row]
                    code_idx = -1
                    for idx, val in enumerate(row_strs):
                        if re.match(r"^\d{6}$", val):
                            code_idx = idx
                            break
                    if code_idx != -1:
                        code = row_strs[code_idx]
                        cat_candidates = [row_strs[i] for i in range(code_idx) if row_strs[i]]
                        if cat_candidates:
                            curr_cat = cat_candidates[-1]
                        desc = row_strs[code_idx + 1] if len(row_strs) > code_idx + 1 else ""
                        points = row_strs[code_idx + 2] if len(row_strs) > code_idx + 2 else ""
                        days = row_strs[code_idx + 3] if len(row_strs) > code_idx + 3 else ""
                        activities.append({
                            "activity_code": code,
                            "category": curr_cat,
                            "activity_name": desc,
                            "points": points,
                            "converted_days": days,
                            "formula": "Điểm / 10 = Ngày CTXH"
                        })
                    else:
                        text_line = " ".join(r for r in row_strs if r)
                        if len(text_line) > 15:
                            _extract_rules_from_text(text_line, calculation_rules)
            wb.close()
        except Exception as e:
            sys.stderr.write(f"[sync_ctxh_rules] Loi doc xlsx {fpath.name}: {e}\n")

    return {
        "title": "QUY ĐỊNH MÃ HÓA, THANG ĐIỂM & QUY ĐỔI NGÀY CTXH (HCMUTE)",
        "source_file": fpath.name,
        "calculation_rules": calculation_rules,
        "total_activities": len(activities),
        "activities": activities
    }


def sync_ctxh_rules_from_file() -> dict:
    """
    Boc tach du lieu hoat dong CTXH va cach tinh diem tu cac file Excel (.xlsx), Word (.docx), PDF, Text (.txt) trong data/.
    Bo hoan toan viec boc tach tu cac file JSON ma thay bang Excel va cac van ban quy che.
    Ghi ket qua ra data/activities.json va data/ctxh.json.
    """
    rule_files = _find_all_ctxh_rules_files()

    merged_activities: dict[str, dict] = {}
    merged_rules = {
        "formula": "Lấy tổng điểm hoạt động CTXH SV đã tích lũy chia cho 10 (số lẻ dưới 5 làm tròn xuống; tròn 5 điểm, quy đổi thành nửa ngày CTXH; trên 5 làm tròn thành 10, quy đổi thành một ngày CTXH). Quy định 2025 áp dụng chia cho 12 đối với các hoạt động chuẩn mới.",
        "deduction_rules": [
            "SV đăng ký tham gia hoạt động CTXH nhưng không thực hiện mà không có lý do hoặc lý do không chính đáng bị trừ 0,5 đến 2,0 ngày CTXH và hạ bậc ĐRL.",
            "SV không hoàn thành nhiệm vụ, gây ảnh hưởng xấu bị trừ từ 01 đến toàn bộ số ngày đã tích lũy.",
            "SV gian lận công nghệ (Fake GPS, can thiệp camera điểm danh) bị hủy bỏ toàn bộ số ngày CTXH đã tích lũy và xử lý kỷ luật."
        ],
        "exclusions": [
            "Các hoạt động không mang tính chất CTXH như xem biểu diễn văn nghệ, xem thể thao, du lịch, tham quan, sinh hoạt CLB không được tính ngày CTXH.",
            "Khi SV tham gia hoạt động CTXH đã được trả tiền công hoặc trả phí thì không được tính ngày CTXH."
        ]
    }

    scanned_files: list[str] = []
    for fpath in rule_files:
        try:
            parsed = _parse_ctxh_rules_from_file(fpath)
            scanned_files.append(fpath.name)
            for act in parsed.get("activities", []):
                code = str(act.get("activity_code", "")).strip()
                if not code:
                    continue
                if code not in merged_activities:
                    merged_activities[code] = act
                else:
                    for k, v in act.items():
                        if v and not merged_activities[code].get(k):
                            merged_activities[code][k] = v

            calc = parsed.get("calculation_rules", {})
            for r in calc.get("deduction_rules", []):
                if r not in merged_rules["deduction_rules"]:
                    merged_rules["deduction_rules"].append(r)
            for e in calc.get("exclusions", []):
                if e not in merged_rules["exclusions"]:
                    merged_rules["exclusions"].append(e)
            if calc.get("formula") and not merged_rules.get("formula"):
                merged_rules["formula"] = calc["formula"]
        except Exception as e:
            sys.stderr.write(f"[sync_ctxh_rules] Loi khi boc tach file {fpath.name}: {e}\n")

    activities_list = sorted(list(merged_activities.values()), key=lambda x: str(x.get("activity_code", "")))

    if not activities_list:
        msg = "Khong tim thay hoac khong boc tach duoc hoat dong CTXH nao tu data/."
        sys.stderr.write(f"[sync_ctxh_rules] {msg}\n")
        return {"status": "empty", "message": msg, "total_activities": 0}

    final_data = {
        "title": "QUY ĐỊNH MÃ HÓA, THANG ĐIỂM & QUY ĐỔI NGÀY CTXH (HCMUTE)",
        "source_files": scanned_files,
        "calculation_rules": merged_rules,
        "total_activities": len(activities_list),
        "activities": activities_list
    }

    # Luu vao ca activities.json va ctxh.json de dong nhat
    for p in ACTIVITIES_RULES_PATHS:
        p.parent.mkdir(parents=True, exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(final_data, f, ensure_ascii=False, indent=2)

    msg = (
        f"Da boc tach thanh cong {len(activities_list)} hoat dong va cach tinh diem "
        f"tu cac file doc/pdf/txt vao activities.json."
    )
    sys.stderr.write(f"[sync_ctxh_rules] {msg}\n")
    return {
        "status": "success",
        "file": ", ".join(scanned_files[:5]),
        "total_activities": len(activities_list),
        "target_file": "activities.json"
    }


# Initialize and sync store on startup
_ensure_index_loaded()


# ==========================================
# MCP TOOLS (Declared for AI Agents to call)
# ==========================================

@mcp.tool()
def search_ctxh_rules(
    query: str,
    top_k: int = 3,
    mode: str = "hybrid",
    year: Optional[str] = None,
) -> str:
    """Tra cuu van ban quy dinh Cong tac xa hoi (CTXH) bang Hybrid Search (BM25 + Dense Vector).
    Tra ve cac dieu khoan, ma hoat dong (6 chu so), thang diem, cach quy doi ngay CTXH, va dieu kien tot nghiep.
    Bat buoc goi cong cu nay khi tra cuu ma hoat dong, quy dinh ngay CTXH hoac tieu chuan mien giam.

    Args:
        query: Tu khoa, ma hoat dong (001101, 001202...), hoac cau hoi can tra cuu quy dinh CTXH.
        top_k: So luong trich doan can lay (mac dinh 3, toi da 6).
        mode: Che do tim kiem ('hybrid', 'vector', 'bm25').
        year: Nam ban hanh quy dinh cu the (vi du '2023', '2024', '2025'). Neu de trong, he thong tu dong uu tien ban moi nhat.
    """
    top_k = max(1, min(top_k, 6))
    results = STORE.search(query=query, top_k=top_k, mode=mode, year=year)

    if not results:
        return "Khong tim thay quy dinh phu hop trong co so du lieu CTXH."

    # Loc lai lan nua: voi cac chunk cung doc_id, chi giu version lon nhat
    if not year:
        from src.retrieval.vector_store import parse_version_key
        max_v_in_results: dict[str, tuple] = {}
        for chunk, _ in results:
            d_id = chunk.doc_id or (chunk.metadata or {}).get("doc_id", "")
            vk = parse_version_key((chunk.metadata or {}).get("version"))
            if d_id not in max_v_in_results or vk > max_v_in_results[d_id]:
                max_v_in_results[d_id] = vk

        results = [
            (chunk, score) for chunk, score in results
            if parse_version_key((chunk.metadata or {}).get("version")) == max_v_in_results.get(
                chunk.doc_id or (chunk.metadata or {}).get("doc_id", "")
            )
        ]

    formatted_passages = []
    for rank, (chunk, score) in enumerate(results, 1):
        meta = chunk.metadata or {}
        file_name = meta.get("file_name", "unknown")
        version = meta.get("version") or "Chưa xác định"
        page_str = f" | Trang {meta.get('page_number')}" if "page_number" in meta else ""

        passage = (
            f"--- [Trich doan {rank}] ---\n"
            f"[Nguồn: {file_name}{page_str} | Phiên bản: {version} | Chunk ID: {chunk.chunk_id}]\n"
            f"Do tuong dong: {score:.4f}\n"
            f"Noi dung:\n{chunk.content}"
        )
        formatted_passages.append(passage)

    return "\n\n".join(formatted_passages)


# Alias for backward compatibility if needed
@mcp.tool()
def search_regulations(query: str, top_k: int = 4, year: Optional[str] = None) -> str:
    """Tra cuu quy dinh CTXH. Tuong duong voi search_ctxh_rules."""
    return search_ctxh_rules(query=query, top_k=top_k, mode="hybrid", year=year)


@mcp.tool()
def list_indexed_documents() -> str:
    """Liet ke toan bo cac van ban quy dinh ve CTXH hien co trong he thong."""
    if not STORE.chunks:
        return "Co so du lieu tai lieu CTXH hien dang trong."

    file_summary: dict[str, dict] = {}
    for chunk in STORE.chunks.values():
        meta = chunk.metadata or {}
        fname = meta.get("file_name", chunk.doc_id)
        if fname not in file_summary:
            file_summary[fname] = {
                "doc_id": chunk.doc_id,
                "version": meta.get("version", ""),
                "extension": meta.get("extension", ""),
                "metadata_source": meta.get("metadata_source", "none"),
                "chunk_count": 0,
                "total_pages": meta.get("total_pages"),
            }
        file_summary[fname]["chunk_count"] += 1

    lines = ["Danh muc van ban quy che CTXH hien co:"]
    for fname, info in file_summary.items():
        ver_str = f", Phien ban: {info['version']}" if info['version'] else ""
        source_str = f", Nguon metadata: {info['metadata_source']}"
        pages_info = f", So trang: {info['total_pages']}" if info["total_pages"] else ""
        lines.append(
            f"  - [{info['doc_id']}] {fname} (Dinh dang: {info['extension']}{ver_str}{source_str}, Chunks: {info['chunk_count']}{pages_info})"
        )

    return "\n".join(lines)


@mcp.tool()
def get_document_full(doc_id: str) -> str:
    """Trich xuat toan van mot van ban quy dinh CTXH theo ma doc_id."""
    matched_chunks = sorted(
        [c for c in STORE.chunks.values() if c.doc_id == doc_id],
        key=lambda c: c.chunk_index,
    )

    if not matched_chunks:
        return f"Khong tim thay van ban nao voi ma doc_id '{doc_id}'."

    meta = matched_chunks[0].metadata or {}
    file_name = meta.get("file_name", doc_id)

    full_text = "\n\n".join(c.content for c in matched_chunks)
    if len(full_text) > 4000:
        full_text = full_text[:4000] + "\n[ ... van ban da duoc cat ngan de dam bao an toan ngu canh ... ]"

    return f"=== Toan van van ban: {file_name} ({doc_id}) ===\n\n{full_text}"


@mcp.tool()
def export_ctxh_rules_json() -> str:
    try:
        result = sync_ctxh_rules_from_file()
    except Exception as e:
        return f"Loi khi boc tach quy che CTXH: {e}"

    if result["status"] == "success":
        return (
            f"[CTXH RULES - THANH CONG] Da boc tach du lieu tu cac file '{result['file']}' vao '{result['target_file']}'.\n"
            f"  - Tong so hoat dong quy doi: {result['total_activities']}\n"
            f"  - Bao gom: Ma hoat dong (6 chu so), linh vuc, thang diem, quy doi ngay va quy tac tinh diem.\n"
            f"  - File ket qua: data/{result['target_file']}"
        )
    elif result["status"] == "empty":
        return f"[CTXH RULES - TRONG] {result['message']}"
    else:
        return f"[CTXH RULES - LOI] {result.get('message', 'Loi khong xac dinh.')}"


if __name__ == "__main__":
    mcp.run(transport="stdio")
