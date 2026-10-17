import json
import os
import shutil
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from rag_server import (
    get_activity,
    register_ctxh_activity,
    sync_students_from_excel,
    convert_points_to_ctxh_days,
    STORE,
    DATA_DIR,
)
from src.evaluation.faithfulness import (
    check_answer_relevance,
    check_context_precision,
    FaithfulnessEvaluator,
)

def run_tests():
    act_path = os.path.join(DATA_DIR, "student_activities.json")
    act_bak = act_path + ".bak_test"
    shutil.copy2(act_path, act_bak)
    try:
        print("=== TEST 1: Football query (not in rules) ===")
        res1 = get_activity("Cổ vũ giải bóng đá Ngoại hạng Anh tại rạp chiếu phim")
        print("Res 1:", res1)
        assert res1.startswith("Error:"), "Test 1 failed: Expected Error: prefix"

        print("\n=== TEST 2: Hien mau query ===")
        res2 = get_activity("hiến máu")
        print("Res 2:\n", res2)
        assert "001201" in res2 and "001202" in res2, "Test 2 failed: Should contain 001201 and 001202"
        assert "Tim thay 2 hoat dong" in res2, "Test 2 failed: Count should be 2"

        print("\n=== TEST 3: Register 001201 with points=10 ===")
        res3 = register_ctxh_activity("23110005", "001201", points=10, note="Test points 10")
        print("Res 3:\n", res3)
        assert "Diem ghi nhan: 10" in res3, "Test 3 failed: points should be 10"
        assert "+1.0 ngay" in res3, "Test 3 failed: days should be +1.0 ngay"

        print("\n=== TEST 4: Register 001101 without points ===")
        res4 = register_ctxh_activity("23110234", "001101", note="Test default points")
        print("Res 4:\n", res4)
        assert "Diem ghi nhan: 25" in res4, "Test 4 failed: default points should be 25"
        assert "+2.5 ngay" in res4, "Test 4 failed: default days should be +2.5 ngay"
        assert "mức cơ sở" in res4.lower() or "muc co so" in res4.lower() or "toi thieu" in res4.lower() or "tối thiểu" in res4.lower(), "Test 4 failed: should mention co so/toi thieu"

        print("\n=== TEST 5: Missing students.xlsx sync ===")
        xlsx_path = os.path.join(DATA_DIR, "students.xlsx")
        bak_path = os.path.join(DATA_DIR, "students.xlsx.bak")
        json_path = os.path.join(DATA_DIR, "students.json")
        with open(json_path, "r", encoding="utf-8") as f:
            before_students = f.read()

        if os.path.exists(xlsx_path):
            shutil.move(xlsx_path, bak_path)
        try:
            res5 = sync_students_from_excel()
            print("Res 5:", res5)
            with open(json_path, "r", encoding="utf-8") as f:
                after_students = f.read()
            assert before_students == after_students, "Test 5 failed: students.json was modified"
            data5 = res5 if isinstance(res5, dict) else json.loads(res5)
            assert data5.get("status") == "kept", "Test 5 failed: status should be 'kept'"
        finally:
            if os.path.exists(bak_path):
                shutil.move(bak_path, xlsx_path)

        print("\n=== TEST 6: Real Chunk Citation Verification ===")
        res_code = get_activity("001202")
        print("Res 001202:\n", res_code)
        assert "Chunk ID:" in res_code, "Test 6 failed: Must contain Chunk ID"
        import re
        m = re.search(r"Chunk ID:\s*([^\s\]]+)", res_code)
        assert m, "Test 6 failed: Cannot parse Chunk ID from citation"
        cid = m.group(1)
        assert cid in STORE.chunks, f"Test 6 failed: Chunk ID '{cid}' does not exist in STORE.chunks"
        chunk_obj = STORE.chunks[cid]
        assert "001202" in chunk_obj.content, f"Test 6 failed: Chunk content does not contain 001202"
        print(f"Verified Real Chunk: {cid} in {chunk_obj.metadata.get('file_name')} (Version: {chunk_obj.metadata.get('version')})")

        print("\n=== TEST 7: Natural language query check ===")
        res_nat = get_activity("mã hoạt động trực tiếp hiến máu là gì")
        print("Res natural language query:\n", res_nat)
        assert "001202" in res_nat, "Test 7 failed: Should match 001202 for natural query"

        print("\n=== TEST 8: CTXH points rounding calculation ===")
        assert convert_points_to_ctxh_days(17) == 2.0, "17 pts should convert to 2.0 days"
        assert convert_points_to_ctxh_days(14) == 1.0, "14 pts should convert to 1.0 day"
        assert convert_points_to_ctxh_days(15) == 1.5, "15 pts should convert to 1.5 days"
        assert convert_points_to_ctxh_days(35) == 3.5, "35 pts should convert to 3.5 days"
        print("CTXH rounding formula verified: 17 pts -> 2.0 days, 14 pts -> 1.0 day, 35 pts -> 3.5 days")

        print("\n=== TEST 9: Evaluation Judge None handling (No fake 1.0) ===")
        rel_empty = check_answer_relevance("", "")
        assert rel_empty is None, f"Expected None on empty relevance, got {rel_empty}"
        prec_empty = check_context_precision("", [])
        assert prec_empty is None, f"Expected None on empty precision, got {prec_empty}"
        evaluator = FaithfulnessEvaluator(groq_client=None)
        rep_empty = evaluator.evaluate("xin chào bạn", "")
        assert rep_empty.score is None, f"Expected None score on no claims/empty context, got {rep_empty.score}"
        print("Evaluation Judge None handling verified: No fake 1.0 returned on error or unjudged cases")

        print("\nALL VERIFICATION TESTS PASSED SUCCESSFULLY!")
    finally:
        if os.path.exists(act_bak):
            shutil.move(act_bak, act_path)

if __name__ == "__main__":
    run_tests()
