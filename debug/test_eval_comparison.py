import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from rag_server import STORE
from tests.evaluate_testset import run_configuration_comparison

eval_dir = _ROOT / "logs" / "eval"
charts_dir = eval_dir / "charts"
eval_dir.mkdir(parents=True, exist_ok=True)
charts_dir.mkdir(parents=True, exist_ok=True)

print("Running configuration comparison test...")
records = run_configuration_comparison(STORE, eval_dir, charts_dir)
print(f"Comparison completed successfully! {len(records)} testcases processed.")
