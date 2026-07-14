"""
generate_examples.py — Regenerate example .pptx files from .md outlines.

Uses placeholder company info (not real BJXH contact data) so the
generated .pptx files are safe for public distribution.

Usage:
    python scripts/generate_examples.py <template.pptx> [--outdir examples/]

Requires the master template .pptx to exist at the given path.
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

DEMO_CONFIG: dict = {
    "company": {
        "company_name": "示例会计师事务所",
        "company_name_en": "Example CPA Firm",
        "phone": "010-XXXX-XXXX",
        "fax": "010-XXXX-XXXX",
        "address": "北京市XX区XX路XX号",
    }
}


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 1

    template_path = Path(sys.argv[1])
    if not template_path.exists():
        print(f"ERROR: template not found: {template_path}", file=sys.stderr)
        return 1

    outdir = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("examples")
    outdir.mkdir(parents=True, exist_ok=True)

    # Write temporary demo config
    demo_cfg_path = outdir / ".demo-config.json"
    demo_cfg_path.write_text(
        json.dumps(DEMO_CONFIG, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Wrote demo config: {demo_cfg_path}")

    # Find all .md outlines in examples/
    outlines = sorted(outdir.glob("*.md"))
    if not outlines:
        print("No .md outlines found in examples/", file=sys.stderr)
        demo_cfg_path.unlink()
        return 1

    from_outline = str(Path(__file__).parent / "from_outline.py")

    for md_path in outlines:
        out_name = md_path.stem + ".pptx"
        out_path = outdir / out_name

        print(f"Generating: {md_path.name} → {out_name} ...")
        result = subprocess.run(
            [
                sys.executable, from_outline,
                str(md_path), str(out_path), str(template_path),
                "--no-chapter-covers",
            ],
            capture_output=True, text=True,
            env={**__import__("os").environ, "BJXH_TEMPLATE": str(template_path)},
        )
        if result.returncode != 0:
            print(f"  FAILED: {result.stderr.strip()}", file=sys.stderr)
        else:
            print(f"  OK: {out_name}")

    demo_cfg_path.unlink(missing_ok=True)
    print("Done. Run python scripts/qa.py examples/<file>.pptx to verify.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
