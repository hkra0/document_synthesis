"""Rerun historical probes, recording the now-expected integrity rejection."""
import importlib.util
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO))
from lib.content_integrity import ContentIntegrityError

spec = importlib.util.spec_from_file_location(
    "historical_probes", REPO / "docs/acceptance/review_n0_n10_probes.py"
)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
verify = module.verify_content_integrity


def record_integrity_result(*args, **kwargs):
    try:
        return {"accepted": verify(*args, **kwargs)}
    except ContentIntegrityError as exc:
        return {"accepted": False, "exception": type(exc).__name__, "reason": str(exc)}


if __name__ == "__main__":
    # The actual production verifier runs; only its expected exception is recorded.
    module.verify_content_integrity = record_integrity_result
    sys.argv = [str(spec.origin), str(REPO / "output/usability-20260908/probes"),
                str(Path(__file__).with_name("independent-probes.json"))]
    module.main()
