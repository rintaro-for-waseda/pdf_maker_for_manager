import ast
from pathlib import Path


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "index.py"
OUT_DIR = ROOT / "docs"
OUT_FILE = OUT_DIR / "index.html"


def extract_page_constant(source_path):
    tree = ast.parse(source_path.read_text(encoding="utf-8"), filename=str(source_path))
    for node in tree.body:
        if isinstance(node, ast.Assign):
            names = [target.id for target in node.targets if isinstance(target, ast.Name)]
            if "PAGE" in names and isinstance(node.value, ast.Constant):
                return node.value.value
    raise RuntimeError("PAGE constant was not found in index.py")


def main():
    html = extract_page_constant(SOURCE)
    OUT_DIR.mkdir(exist_ok=True)
    OUT_FILE.write_text(html, encoding="utf-8")
    (OUT_DIR / ".nojekyll").write_text("", encoding="utf-8")
    print(f"wrote {OUT_FILE.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
