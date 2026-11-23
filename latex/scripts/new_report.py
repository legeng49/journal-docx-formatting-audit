"""新建一份课程报告：把 assets/template 复制到目标文件夹，并建好 figures/。

用法：python new_report.py <目标文件夹>
目标文件夹不存在时会创建；已存在且非空时报错退出，避免覆盖已有报告。
"""
import shutil
import sys
from pathlib import Path

TEMPLATE = Path(__file__).resolve().parent.parent / "assets" / "template"


def main() -> None:
    if len(sys.argv) != 2:
        sys.exit("用法：python new_report.py <目标文件夹>")
    dest = Path(sys.argv[1]).resolve()
    if dest.exists() and any(dest.iterdir()):
        sys.exit(f"目标文件夹已存在且非空：{dest}")
    dest.mkdir(parents=True, exist_ok=True)
    for f in TEMPLATE.iterdir():
        shutil.copy2(f, dest / f.name)
    (dest / "figures").mkdir()
    print(f"已创建：{dest}")
    for f in sorted(dest.iterdir()):
        print("  ", f.name)


if __name__ == "__main__":
    main()
