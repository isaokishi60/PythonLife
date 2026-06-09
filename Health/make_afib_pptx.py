from pathlib import Path
from pptx import Presentation
from pptx.util import Inches

# ユーザー名 spax2 / kawgu の違いを自動吸収
BASE_DIR = Path.home() / "OneDrive" / "ドキュメント" / "PythonWork" / "Health"

PERIOD_PNG_DIR = BASE_DIR / "01_Garmin_Import" / "outputs_period" / "png"
NIGHT_PNG_DIR = BASE_DIR / "01_Garmin_Import" / "outputs" / "png"

OUT_PPTX = BASE_DIR / "心房細動グラフ.pptx"


def latest_file(folder: Path, pattern: str) -> Path:
    files = list(folder.glob(pattern))
    if not files:
        raise FileNotFoundError(f"ファイルが見つかりません: {folder}\\{pattern}")
    return max(files, key=lambda p: p.stat().st_mtime)


slides = [
    ("DailyBeats", latest_file(PERIOD_PNG_DIR, "HeartPeriod_DailyBeats_2025-10-01_*.png")),
    ("RHR",        latest_file(PERIOD_PNG_DIR, "HeartPeriod_RHR_2025-10-01_*.png")),
    ("Tachy",      latest_file(PERIOD_PNG_DIR, "HeartPeriod_Tachy_2025-10-01_*.png")),
    ("NightHR",    latest_file(NIGHT_PNG_DIR, "RestHR_Night_*_21-06.png")),
]

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

for title, png in slides:
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    text_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.15), Inches(12.5), Inches(0.4))
    text_box.text = title

    slide.shapes.add_picture(
        str(png),
        Inches(0.4),
        Inches(0.65),
        width=Inches(12.5)
    )

    print(f"{title}: {png.name}")

prs.save(OUT_PPTX)
print(f"Saved: {OUT_PPTX}")
print(f"スライド数: {len(slides)}")