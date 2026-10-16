import os
import sys
import json
import uuid
from pathlib import Path
from datetime import datetime
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

from src.loaders import DocumentLoaderFactory, SUPPORTED_EXTENSIONS
from src.chunking import RecursiveChunker
from src.retrieval import VectorStore, EmbeddingGenerator
from src.models import StudentCTXHRecord, CTXHRegistration

load_dotenv()

# Initialize MCP Server
mcp = MCPServer("rag-knowledge-server")

DATA_DIR = _HERE / "data"
INDEX_PATH = DATA_DIR / "index.json"
CTXH_DB_PATH = DATA_DIR / "student_ctxh.json"
STUDENTS_DB_PATH = DATA_DIR / "students.json"

# Global Store
STORE = VectorStore()


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
            if is_modified:
                STORE.remove_chunks_by_file(fname)
                updated_files.append(fname)
            else:
                added_files.append(fname)

            docs = DocumentLoaderFactory.load_file(fpath)
            chunks = chunker.split_documents(docs)

            if chunks:
                if emb_gen is None:
                    emb_gen = EmbeddingGenerator()
                texts = [c.content for c in chunks]
                embeddings = emb_gen.embed_documents(texts, batch_size=16)
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
    """Load cached index and incrementally synchronize with data/."""
    global STORE
    if INDEX_PATH.exists():
        try:
            STORE = VectorStore.load_from_json(INDEX_PATH)
        except Exception:
            STORE = VectorStore()

    sync_knowledge_base()


# Database helper functions for CTXH records
def _load_student_ctxh() -> list[dict]:
    if not CTXH_DB_PATH.exists():
        return []
    try:
        with open(CTXH_DB_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def _save_student_ctxh(data: list[dict]) -> None:
    CTXH_DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(CTXH_DB_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _load_students() -> list[dict]:
    if not STUDENTS_DB_PATH.exists():
        return []
    try:
        with open(STUDENTS_DB_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


# Initialize and sync store on startup
_ensure_index_loaded()


if __name__ == "__main__":
    mcp.run(transport="stdio")