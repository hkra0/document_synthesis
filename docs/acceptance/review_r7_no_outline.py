"""Re-evaluate copies of R7 samples with heading metadata removed.

Usage: python3 docs/acceptance/review_r7_no_outline.py SCRATCH_DIR METRICS_JSON
Original samples and labels remain unchanged; visual direct formatting is retained.
"""
import importlib.util
import json
import sys
from pathlib import Path

from docx import Document

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))


def main():
    samples = Path(sys.argv[1]).resolve()
    samples.mkdir(parents=True, exist_ok=True)
    evaluation = ROOT / "docs/acceptance/r7-evaluation"
    dataset = json.loads((evaluation / "dataset.json").read_text())
    for item in dataset["documents"]:
        doc = Document(evaluation / "samples" / item["filename"])
        for index in item["heading_paragraph_indices"]:
            paragraph = doc.paragraphs[index]
            paragraph.style = "Normal"
            for outline in paragraph._p.xpath("./w:pPr/w:outlineLvl"):
                outline.getparent().remove(outline)
        doc.save(samples / item["filename"])
    spec = importlib.util.spec_from_file_location("r7_eval", evaluation / "evaluate_dataset.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    metrics = module.evaluate_dataset(dataset, samples)
    metrics["audit_transformation"] = (
        "Copied existing synthetic evaluation samples; only labeled heading paragraphs "
        "were changed to Normal and their direct outlineLvl removed. Direct visual "
        "formatting, text, paragraph positions and original ground truth are unchanged."
    )
    Path(sys.argv[2]).write_text(json.dumps(metrics, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps(metrics["overall"], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
