#!/usr/bin/env python3
"""
Figma PPT Insert - Marker Image Insertion

Insert Figma-exported screen images into a PowerPoint template. Targets are
found per slide, in this order:
  1. marker shapes whose text is ``[화면 교체 위치: {화면ID}]`` (or pictures this
     tool inserted earlier, which keep their screen ID and original box);
  2. otherwise, a table cell "화면ID" gives the ID and the slide's largest
     picture (the existing screenshot) is replaced in its box and z-order.
Each image is matched by file name (= Figma frame name = screen ID) and fitted
into the target box; the marker / old picture is removed.

All geometry is computed in inches; image pixels are used only for the
unitless aspect ratio.

Project layout (one folder per project under <repo>/projects/):
    projects/{project}/01_template/                one source .pptx
    projects/{project}/02_images/{ID}.png          one current image per screen ID
    projects/{project}/02_images/_archive/         replaced images ({ID}_{YYYYMMDD_HHMM}.png)
    projects/{project}/03_output/{name}_최종.pptx  latest result
    projects/{project}/03_output/_archive/         previous results (newest 10 kept)
    projects/{project}/insert_log.md               cumulative run log

Before matching, exports downloaded today (names shaped like the deck's screen
IDs, or Figma zip exports of them) are moved from ~/Downloads into 02_images/;
the image each one replaces goes to 02_images/_archive/. Matching files from the
two days before are only reported. Date folders from the earlier layout
(02_images/YYYY-MM-DD/) are folded in the same way.

Usage:
    python3 .claude/skills/figma-ppt-insert/scripts/insert.py [options]

Examples:
    python3 .claude/skills/figma-ppt-insert/scripts/insert.py --new-project IMS
    python3 .claude/skills/figma-ppt-insert/scripts/insert.py --project IMS
    python3 .claude/skills/figma-ppt-insert/scripts/insert.py --project IMS --mode fill
    python3 .claude/skills/figma-ppt-insert/scripts/insert.py --project IMS --dry-run
    python3 .claude/skills/figma-ppt-insert/scripts/insert.py --project IMS --download-days 3
    python3 .claude/skills/figma-ppt-insert/scripts/insert.py --list

Dependencies:
    python-pptx, Pillow
"""
from __future__ import annotations

import argparse
import copy
import re
import shutil
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Iterator, Optional

# scripts/insert.py: [0]=scripts [1]=figma-ppt-insert [2]=skills [3]=.claude [4]=repo
REPO_ROOT = Path(__file__).resolve().parents[4]

try:
    from PIL import Image
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from pptx.oxml.ns import qn
    from pptx.util import Emu, Inches, Pt
except ImportError as exc:  # pragma: no cover - environment guard
    print(
        f"Missing dependency: {exc.name}. Install with: pip install python-pptx Pillow",
        file=sys.stderr,
    )
    raise SystemExit(1)

PROJECTS_DIR = REPO_ROOT / "projects"
TEMPLATE_DIR = "01_template"
IMAGES_DIR = "02_images"
OUTPUT_DIR = "03_output"
ARCHIVE_DIR = "_archive"
LOG_FILE = "insert_log.md"
OUTPUT_SUFFIX = "_최종"
ARCHIVE_KEEP = 10
IMAGE_EXTS = (".png", ".jpg", ".jpeg")  # earlier = preferred on a same-time tie
IMAGE_FORMATS = {"PNG", "JPEG"}  # actual content, checked with Pillow
IGNORED_FILES = {".gitkeep", "thumbs.db", "desktop.ini", ".ds_store"}
OVERLAY_KEEP = "keep"  # shapes drawn over the old screenshot stay as they are
OVERLAY_CLEAN = "clean"  # remove them, keeping callout badges, dashed frames and arrows
BADGE_MAX_IN = 0.5  # callout badges are small shapes holding just a number such as 1, 2a, 1-1
OVERLAY_MARGIN_IN = 0.05
HEADER_GAP_IN = 0.1  # a header bar ending this close to the screen top belongs to the screen
HEADER_MIN_SHARE = 0.9  # ...when it spans at least this share of the screen width  # a shape counts as "on the screen" when inside the box plus this margin
MODE_FIT = "fit"
MODE_FILL = "fill"
# Font for text this tool adds to slides (install-local font lock).
FONT_FAMILY = "Pretendard"

TABLE_ID_LABEL = "화면ID"  # compared with whitespace removed, upper-cased
MIN_SCREEN_AREA_SQIN = 1.0  # smaller pictures (logos, icons) are never screen targets
SIMILAR_RATIO_TOLERANCE = 0.15  # e.g. 16:9 vs 16:10 (11%) counts as similar
PLACE_TOP = "상단 맞춤"
PLACE_CENTER = "가운데"
PLACE_FILL = "채우기"
MIN_PPI = 150  # below this, screen text looks blurry when shown on the slide
PAGE_LABEL = "페이지"
UNKNOWN_FIELD_LABELS = {"화면명", "화면타입", "유형", "화면경로"}  # left as placeholders on new pages
PLACEHOLDER = "(확인 필요)"
TYPE_LABEL = "유형"
CLONE_TYPE_VALUE = "변경"  # new pages copy the frame of the nearest slide with this 유형
DESCRIPTION_PLACEHOLDER = (
    "[확인 필요] 화면 설명",
    "피그마 export로 추가된 신규 화면입니다.",
    "화면명·화면타입·유형·화면 경로와 기능 설명을 입력하세요.",
)
SMALL_SCREEN_SHARE = 0.7  # target picture below this share of the median screen area is flagged
HEADER_TAG_MAX_CHARS = 5  # header text this short ("1.9", "9/12") is a tag kept on new pages
HEADER_ZONE = 0.15  # an ID-only text box above this share of the slide height names the screen
FOOTER_ZONE = 0.9  # shapes below this share of the slide height are footer (page number)
FRAME_MIN_SHARE = 0.4  # empty rectangle covering this share of the slide = content frame
DOWNLOADS_DIR = Path.home() / "Downloads"
DOWNLOAD_DAYS = 1  # import exports downloaded today (1) or within the last N days
NOTICE_DAYS = 2  # older downloads within this many extra days are only reported
IMAGE_ARCHIVE_DIR = "_archive"  # replaced images and imported zip exports
PIC_NAME_PREFIX = "FIGMA:"
REGION_TAG = "figma-ppt-insert region_in="

_MARKER_RE = re.compile(r"\[\s*화면\s*교체\s*위치\s*[:：]\s*(?P<id>[^\]\s]+)\s*\]")
_DATE_DIR_RE = re.compile(r"\d{4}-\d{2}-\d{2}")  # legacy date-folder layout
_SCALE_SUFFIX_RE = re.compile(r"@\d+(?:\.\d+)?x$", re.IGNORECASE)
_COPY_SUFFIX_RE = re.compile(r"\s*\(\d+\)$")  # browser duplicate: "UI-IMS-2001 (1).png"
_TRAILING_DIGITS_RE = re.compile(r"\d+$")
_ID_TEXT_RE = re.compile(r"[A-Za-z][A-Za-z0-9]*(?:-[A-Za-z0-9]+){2,}")  # e.g. U-CAS-KVDS-AI-01
_BADGE_TEXT_RE = re.compile(r"\d{1,2}(?:[a-zA-Z]|-\d{1,2})?")
_OCCURRENCE_RE = re.compile(r"^(?P<base>.+)_(?P<n>\d+)$")  # {ID}_{n}: nth slide with that ID
_BAD_PROJECT_CHARS_RE = re.compile(r'[\\/:*?"<>|]')
_NUMBER_SPLIT_RE = re.compile(r"(\d+)")
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REGION_RE = re.compile(re.escape(REGION_TAG) + r"([-\d.]+),([-\d.]+),([-\d.]+),([-\d.]+)")


@dataclass
class Box:
    """Rectangle in inches."""

    left: float
    top: float
    width: float
    height: float


@dataclass
class Target:
    slide_no: int
    screen_id: str
    kind: str  # "marker" | "picture" (inserted earlier) | "table" / "textbox" (화면ID + screenshot)
    box: Box
    element: object  # lxml element of the shape to replace
    occurrence: int = 1  # nth slide carrying this ID, in deck order (matches {ID}_{n} files)


@dataclass
class ImageEntry:
    screen_id: str
    path: Path
    label: str  # file name inside 02_images/


@dataclass
class RunReport:
    inserted: list[tuple[Target, ImageEntry, str]] = field(default_factory=list)
    missing_image: list[Target] = field(default_factory=list)
    kept_existing: list[Target] = field(default_factory=list)
    overlay: dict[int, tuple[int, int]] = field(default_factory=dict)  # slide -> (on screen, badges/frames)
    overlay_mode: str = OVERLAY_KEEP
    header_added: dict[int, float] = field(default_factory=dict)  # slide -> inches added on top
    unused_images: list[ImageEntry] = field(default_factory=list)
    imported: list[tuple[str, str]] = field(default_factory=list)  # (download name, dest)
    moved: list[tuple[str, str]] = field(default_factory=list)  # (old place, new name) in 02_images
    archived: list[tuple[str, str]] = field(default_factory=list)  # (old place, archive name)
    recent_downloads: list[tuple[str, date]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def _key(screen_id: str) -> str:
    return screen_id.strip().upper()


def _emu_to_in(value: int) -> float:
    return Emu(value).inches


def _shape_box(shape) -> Box:
    return Box(
        _emu_to_in(shape.left),
        _emu_to_in(shape.top),
        _emu_to_in(shape.width),
        _emu_to_in(shape.height),
    )


def _marker_id(shape) -> Optional[str]:
    if not getattr(shape, "has_text_frame", False):
        return None
    match = _MARKER_RE.fullmatch(shape.text_frame.text.strip())
    return match.group("id") if match else None


def _iter_group_markers(group) -> Iterator[str]:
    for shape in group.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_group_markers(shape)
        else:
            screen_id = _marker_id(shape)
            if screen_id:
                yield screen_id


def _picture_region(shape) -> Box:
    descr = shape._element.nvPicPr.cNvPr.get("descr", "")
    match = _REGION_RE.search(descr)
    if match:
        return Box(*(float(v) for v in match.groups()))
    return _shape_box(shape)


def _table_screen_id(slide) -> Optional[str]:
    """Return the value next to a "화면ID" header cell in any top-level table."""
    for shape in slide.shapes:
        if not getattr(shape, "has_table", False) or not shape.has_table:
            continue
        for row in shape.table.rows:
            cells = [c.text.strip() for c in row.cells]
            for idx, text in enumerate(cells):
                if re.sub(r"\s+", "", text).upper() != TABLE_ID_LABEL:
                    continue
                # Merged cells repeat or blank out; take the first non-empty value after it.
                for value in cells[idx + 1:]:
                    if value and re.sub(r"\s+", "", value).upper() != TABLE_ID_LABEL:
                        return value.split()[0]
    return None


def _textbox_screen_id(slide, slide_height: float) -> Optional[str]:
    """Screen ID written alone in a header text box (decks without a 화면ID table).

    When several ID-shaped boxes sit in the header (screen ID left, requirement
    ID right), the leftmost one is the screen ID.
    """
    found = []
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False) or shape.top is None:
            continue
        text = shape.text_frame.text.strip()
        if _ID_TEXT_RE.fullmatch(text) and _emu_to_in(shape.top) < slide_height * HEADER_ZONE:
            found.append((shape.left, text))
    return min(found)[1] if found else None


def _slide_screen_id(slide, slide_height: float) -> tuple[Optional[str], str]:
    """(screen ID, "table" | "textbox") for a screen slide without markers."""
    screen_id = _table_screen_id(slide)
    if screen_id:
        return screen_id, "table"
    return _textbox_screen_id(slide, slide_height), "textbox"


def _largest_picture(slide):
    pictures = [
        s for s in slide.shapes
        if s.shape_type == MSO_SHAPE_TYPE.PICTURE
        and _emu_to_in(s.width) * _emu_to_in(s.height) >= MIN_SCREEN_AREA_SQIN
    ]
    return max(pictures, key=lambda s: s.width * s.height, default=None)


def find_targets(prs, report: RunReport) -> list[Target]:
    """Collect targets slide by slide.

    A slide with markers or previously inserted pictures uses those. Otherwise,
    when a table carries a "화면ID" cell (or a header text box holds just the
    ID), the slide's largest picture (the existing screenshot) becomes the
    target and keeps its box and z-order. Each target records which occurrence
    of its ID it is, so ``{ID}_2`` can go to the second slide with that ID.
    """
    targets: list[Target] = []
    slide_height = _emu_to_in(prs.slide_height)
    seen: dict[str, int] = {}  # slides so far per ID, counting ID slides with no picture too

    def count(screen_id: str) -> int:
        seen[_key(screen_id)] = seen.get(_key(screen_id), 0) + 1
        return seen[_key(screen_id)]

    for slide_no, slide in enumerate(prs.slides, start=1):
        before = len(targets)
        has_group_marker = _scan_slide(slide, slide_no, targets, report)
        if has_group_marker or len(targets) > before:
            occurrences = {_key(t.screen_id): 0 for t in targets[before:]}
            for target in targets[before:]:
                key = _key(target.screen_id)
                occurrences[key] = occurrences[key] or count(target.screen_id)
                target.occurrence = occurrences[key]
            continue
        screen_id, kind = _slide_screen_id(slide, slide_height)
        if not screen_id:
            continue
        occurrence = count(screen_id)
        picture = _largest_picture(slide)
        if picture is None:
            report.warnings.append(
                f"슬라이드 {slide_no}: 화면ID `{screen_id}`는 있지만 교체할 화면 이미지가 없습니다 — "
                "마커 도형을 넣어 주세요."
            )
            continue
        targets.append(
            Target(slide_no, screen_id, kind, _shape_box(picture), picture._element, occurrence)
        )
    return targets


def _warn_small_screens(targets: list[Target], replaced: list[Target], report: RunReport) -> None:
    """Flag a "largest picture" that is much smaller than the deck's usual screen.

    On pages drawn mostly with shapes the largest picture can be a chart or a
    heatmap rather than the screen itself.
    """
    picked = [t for t in targets if t.kind in ("table", "textbox")]
    if len(picked) < 3:
        return
    areas = sorted(t.box.width * t.box.height for t in picked)
    usual = areas[len(areas) // 2]
    for target in replaced:
        if target.kind in ("table", "textbox") and target.box.width * target.box.height < usual * SMALL_SCREEN_SHARE:
            report.warnings.append(
                f"슬라이드 {target.slide_no}: 교체 대상 그림({target.box.width:.1f}×{target.box.height:.1f}in)이 "
                "다른 화면보다 작습니다 — 화면 전체가 아니라 차트 등 일부 그림일 수 있으니 결과를 확인하세요."
            )


def _scan_slide(slide, slide_no: int, targets: list[Target], report: RunReport) -> bool:
    """Append marker / inserted-picture targets; return True if a group marker was seen."""
    group_marker = False
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for screen_id in _iter_group_markers(shape):
                group_marker = True
                report.warnings.append(
                    f"슬라이드 {slide_no}: 그룹 안의 마커 `{screen_id}`는 1차 미지원 — "
                    "그룹을 해제한 뒤 다시 실행하세요."
                )
            continue
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and shape.name.startswith(PIC_NAME_PREFIX):
            screen_id = shape.name[len(PIC_NAME_PREFIX):]
            targets.append(
                Target(slide_no, screen_id, "picture", _picture_region(shape), shape._element)
            )
            continue
        screen_id = _marker_id(shape)
        if screen_id:
            if shape.rotation:
                report.warnings.append(
                    f"슬라이드 {slide_no}: 마커 `{screen_id}`가 회전돼 있습니다 — "
                    "회전 없이 영역만 사용합니다."
                )
            targets.append(Target(slide_no, screen_id, "marker", _shape_box(shape), shape._element))
    return group_marker


def _is_ignorable(path: Path) -> bool:
    name = path.name.lower()
    return name in IGNORED_FILES or name.startswith(("~$", "._"))


def _unsupported_reason(path: Path) -> Optional[str]:
    """Return why the file cannot be inserted, or None when it is a usable PNG/JPEG."""
    ext = path.suffix.lower()
    if ext not in IMAGE_EXTS:
        return f"지원하지 않는 형식({ext or '확장자 없음'})"
    try:
        with Image.open(path) as img:
            actual = img.format
    except (OSError, ValueError):
        return "이미지를 열 수 없음(손상된 파일)"
    if actual not in IMAGE_FORMATS:
        return f"확장자는 {ext}이지만 실제 형식은 {actual}"
    return None


def _clean_stem(file_name: str) -> str:
    """Screen ID from a file name: drop ``@2x`` scale and browser `` (1)`` copy suffixes."""
    stem = _COPY_SUFFIX_RE.sub("", Path(file_name).stem)
    return _SCALE_SUFFIX_RE.sub("", stem).strip()


def deck_screen_ids(prs) -> set[str]:
    """Screen IDs the deck already knows (markers, inserted pictures, 화면ID tables)."""
    ids = {_key(t.screen_id) for t in find_targets(prs, RunReport())}
    height = _emu_to_in(prs.slide_height)
    ids.update(_key(sid) for sid, _ in (_slide_screen_id(s, height) for s in prs.slides) if sid)
    return ids


class _IdMatcher:
    """Tell export files from unrelated downloads by the deck's screen-ID shape.

    A name counts when it is a known ID or a known ID prefix followed by digits
    (``UI-IMS-2001`` in the deck lets ``UI-IMS-2402`` in as a new screen).
    """

    def __init__(self, known_ids: set[str]) -> None:
        self.known = known_ids
        self.prefixes = {
            _TRAILING_DIGITS_RE.sub("", sid) for sid in known_ids if _TRAILING_DIGITS_RE.search(sid)
        }
        self.prefixes.discard("")

    def screen_id(self, file_name: str) -> Optional[str]:
        stem = _clean_stem(file_name)
        occurrence = _OCCURRENCE_RE.match(stem)
        base = occurrence.group("base") if occurrence else stem
        return stem if self._is_id(base) else None

    def _is_id(self, text: str) -> bool:
        key = _key(text)
        if key in self.known:
            return True
        return any(
            key.startswith(prefix) and key[len(prefix):].isdigit() for prefix in self.prefixes
        )


def _download_time(path: Path) -> float:
    # Some copies keep the original mtime; creation (birth) time is when it landed here.
    # st_ctime is not used: on recent Windows Pythons it is the metadata-change time.
    stat = path.stat()
    return max(stat.st_mtime, getattr(stat, "st_birthtime", 0.0))


def _download_date(path: Path) -> date:
    return datetime.fromtimestamp(_download_time(path)).date()


def _zip_members(path: Path, matcher: _IdMatcher) -> list[tuple[str, str]]:
    """(member name, screen ID) for image members of a Figma zip export."""
    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
    except (zipfile.BadZipFile, OSError):
        return []
    found = []
    for name in names:
        base = Path(name).name
        if Path(base).suffix.lower() in IMAGE_EXTS:
            screen_id = matcher.screen_id(base)
            if screen_id:
                found.append((name, screen_id))
    return found


@dataclass
class _Candidate:
    """One available export of a screen, wherever it currently lives."""

    screen_id: str
    stamp: float  # export / download time
    path: Path  # the file itself, or the zip holding it
    member: Optional[str]  # zip member name
    origin: str  # ORIGIN_*

    @property
    def ext(self) -> str:
        return Path(self.member or self.path.name).suffix.lower()

    @property
    def name(self) -> str:
        return f"{self.path.name} → {Path(self.member).name}" if self.member else self.path.name


ORIGIN_CURRENT = "current"  # 02_images/{ID}.png, already in place
ORIGIN_LOOSE = "loose"  # 02_images/ file with a suffix to clean (@2x, (1))
ORIGIN_LEGACY = "legacy"  # 02_images/YYYY-MM-DD/ from the earlier date-folder layout
ORIGIN_DOWNLOAD = "download"


def _rank(candidate: _Candidate) -> tuple:
    # Newest wins; on a tie keep the file already in place, then prefer png.
    return (
        candidate.stamp,
        candidate.origin == ORIGIN_CURRENT,
        -IMAGE_EXTS.index(candidate.ext),
    )


def _scan_downloads(
    downloads: Path, matcher: _IdMatcher, report: RunReport, days: int
) -> list[_Candidate]:
    """Exports downloaded within ``days`` days; older matches within NOTICE_DAYS are reported."""
    if not downloads.is_dir():
        report.warnings.append(f"다운로드 폴더를 찾을 수 없어 가져오지 않았습니다: `{downloads}`")
        return []
    today = date.today()
    found: list[_Candidate] = []
    for path in sorted(downloads.iterdir()):
        if not path.is_file() or _is_ignorable(path):
            continue
        ext = path.suffix.lower()
        if ext == ".zip":
            entries = _zip_members(path, matcher)
        elif ext in IMAGE_EXTS:
            sid = matcher.screen_id(path.name)
            entries = [(None, sid)] if sid and not _unsupported_reason(path) else []
        else:
            continue
        if not entries:
            continue
        day = _download_date(path)
        age = (today - day).days
        if age >= days:
            if age < days + NOTICE_DAYS:
                report.recent_downloads.append((path.name, day))
            continue
        stamp = _download_time(path)
        found += [_Candidate(sid, stamp, path, member, ORIGIN_DOWNLOAD) for member, sid in entries]
    return found


def _scan_project(images_root: Path, report: RunReport) -> tuple[list[_Candidate], list[Path]]:
    """Images already in 02_images (flat or legacy date folders) and legacy folders seen."""
    found: list[_Candidate] = []
    legacy_dirs: list[Path] = []
    if not images_root.is_dir():
        return found, legacy_dirs

    def add(path: Path, origin: str, where: str) -> None:
        reason = _unsupported_reason(path)
        if reason:
            report.warnings.append(
                f"`{where}{path.name}` 건너뜀 — {reason}. 피그마에서 PNG(또는 JPG)로 다시 export하세요."
            )
            return
        sid = _clean_stem(path.name)
        if origin == ORIGIN_CURRENT and path.name != f"{sid}{path.suffix}":
            origin = ORIGIN_LOOSE
        found.append(_Candidate(sid, path.stat().st_mtime, path, None, origin))

    for entry in sorted(images_root.iterdir()):
        if entry.is_file():
            if not _is_ignorable(entry):
                add(entry, ORIGIN_CURRENT, "")
        elif entry.name == IMAGE_ARCHIVE_DIR:
            continue
        elif _DATE_DIR_RE.fullmatch(entry.name):
            legacy_dirs.append(entry)
            for path in sorted(entry.iterdir()):
                if path.is_file() and not _is_ignorable(path):
                    add(path, ORIGIN_LEGACY, f"{entry.name}/")
        else:
            report.warnings.append(f"`{IMAGES_DIR}/{entry.name}` 폴더는 사용하지 않아 건너뜁니다.")
    return found, legacy_dirs


def _archive_name(archive_dir: Path, candidate: _Candidate) -> Path:
    stamp = datetime.fromtimestamp(candidate.stamp).strftime("%Y%m%d_%H%M")
    return _unique_path(archive_dir / f"{candidate.screen_id}_{stamp}{candidate.ext}")


def organize_images(
    images_root: Path,
    downloads: list[_Candidate],
    report: RunReport,
    *,
    dry_run: bool,
    scratch: Path,
) -> dict[str, ImageEntry]:
    """Keep exactly one current image per screen ID in 02_images/; archive the rest.

    The newest export of each ID (from Downloads, a leftover date folder, or the
    folder itself) becomes ``02_images/{ID}.{ext}``; the image it replaces moves
    to ``02_images/_archive/{ID}_{YYYYMMDD_HHMM}.{ext}``. Losing downloads stay in
    Downloads. In a dry run nothing moves; zip members are unpacked to ``scratch``.
    """
    project, legacy_dirs = _scan_project(images_root, report)
    by_key: dict[str, list[_Candidate]] = {}
    for candidate in project + downloads:
        by_key.setdefault(_key(candidate.screen_id), []).append(candidate)

    archive_dir = images_root / IMAGE_ARCHIVE_DIR
    images: dict[str, ImageEntry] = {}
    used_zips: set[Path] = set()
    for key in sorted(by_key, key=_natural_key):
        group = sorted(by_key[key], key=_rank, reverse=True)
        winner, losers = group[0], group[1:]
        dest = images_root / f"{winner.screen_id}{winner.ext}"

        for loser in losers:
            if loser.origin == ORIGIN_DOWNLOAD:
                report.warnings.append(
                    f"다운로드의 `{loser.name}`은 같은 화면 ID의 더 최근 이미지가 있어 가져오지 않았습니다."
                )
                continue
            target = _archive_name(archive_dir, loser)
            report.archived.append((_display(loser, images_root), target.name))
            if not dry_run:
                archive_dir.mkdir(parents=True, exist_ok=True)
                shutil.move(str(loser.path), str(target))

        if winner.origin == ORIGIN_DOWNLOAD:
            report.imported.append((winner.name, f"{IMAGES_DIR}/{dest.name}"))
        elif winner.origin != ORIGIN_CURRENT:
            report.moved.append((_display(winner, images_root), dest.name))

        source = winner.path
        if winner.origin != ORIGIN_CURRENT:
            if winner.member:
                used_zips.add(winner.path)
                out = dest if not dry_run else scratch / dest.name
                with zipfile.ZipFile(winner.path) as archive:
                    out.write_bytes(archive.read(winner.member))
                source = out
            elif not dry_run:
                shutil.move(str(winner.path), str(dest))
                source = dest
        elif not dry_run:
            source = dest
        images[key] = ImageEntry(winner.screen_id, source, dest.name)

    if not dry_run:
        for zip_path in sorted(used_zips):
            archive_dir.mkdir(parents=True, exist_ok=True)
            shutil.move(str(zip_path), str(_unique_path(archive_dir / zip_path.name)))
        for folder in legacy_dirs:
            for leftover in folder.rglob("*"):
                if leftover.is_file() and not _is_ignorable(leftover):
                    archive_dir.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(leftover), str(_unique_path(archive_dir / leftover.name)))
            shutil.rmtree(folder, ignore_errors=True)
    if legacy_dirs:
        names = ", ".join(f.name for f in legacy_dirs)
        done = "정리할 예정입니다" if dry_run else "정리했습니다"
        report.warnings.append(
            f"날짜 폴더({names})를 {done} — 화면별 최신 이미지는 `{IMAGES_DIR}/`에, "
            f"이전 이미지는 `{IMAGES_DIR}/{IMAGE_ARCHIVE_DIR}/`에 있습니다."
        )
    return images


def _display(candidate: _Candidate, images_root: Path) -> str:
    try:
        return candidate.path.relative_to(images_root).as_posix()
    except ValueError:
        return candidate.name


def _image_size(path: Path) -> tuple[int, int]:
    with Image.open(path) as img:
        return img.size


def _effective_ppi(size: tuple[int, int], region: Box, placement: str) -> float:
    """Pixels per inch of the image as shown on the slide (quality check only)."""
    per_width, per_height = size[0] / region.width, size[1] / region.height
    if placement == PLACE_TOP:
        return per_width  # width always matches the region
    # center shows the whole image (tighter axis sets the scale); fill covers the box.
    return max(per_width, per_height) if placement == PLACE_CENTER else min(per_width, per_height)


def _is_similar_ratio(ratio: float, region: Box) -> bool:
    return abs(ratio / (region.width / region.height) - 1) <= SIMILAR_RATIO_TOLERANCE


def _top_box(region: Box, ratio: float) -> tuple[Box, float]:
    """Match the region's width and top edge; return the box and bottom crop fraction.

    Top-left and top-right corners coincide with the region, so callouts placed
    on the previous screen keep their position. A slightly taller image is
    cropped at the bottom instead of spilling out of the region.
    """
    height = region.width / ratio
    crop_bottom = 0.0
    if height > region.height:
        crop_bottom = 1 - region.height / height
        height = region.height
    return Box(region.left, region.top, region.width, height), crop_bottom


def _center_box(region: Box, ratio: float) -> Box:
    if ratio > region.width / region.height:
        width, height = region.width, region.width / ratio
    else:
        width, height = region.height * ratio, region.height
    return Box(
        region.left + (region.width - width) / 2,
        region.top + (region.height - height) / 2,
        width,
        height,
    )


def _apply_fill_crop(picture, region: Box, ratio: float) -> None:
    region_ratio = region.width / region.height
    if ratio > region_ratio:
        side = (1 - region_ratio / ratio) / 2
        picture.crop_left = picture.crop_right = side
    elif ratio < region_ratio:
        side = (1 - ratio / region_ratio) / 2
        picture.crop_top = picture.crop_bottom = side


def place_image(slide, target: Target, image: ImageEntry, mode: str) -> tuple[str, float]:
    """Add the picture at the target's z-position and box, then remove the target.

    In fit mode a similar or wider aspect ratio is top-aligned at full width; a
    clearly narrower one (popup, mobile) is centered. Returns the placement used and
    the effective resolution (ppi) of the placed image.
    """
    size = _image_size(image.path)
    ratio = size[0] / size[1]
    region = target.box
    crop_bottom = 0.0
    if mode == MODE_FILL:
        placement, box = PLACE_FILL, region
    elif _is_similar_ratio(ratio, region) or ratio > region.width / region.height:
        # Wider than the box means a full screen, not a popup: keep it on top at full width.
        placement = PLACE_TOP
        box, crop_bottom = _top_box(region, ratio)
    else:
        placement, box = PLACE_CENTER, _center_box(region, ratio)
    picture = slide.shapes.add_picture(
        str(image.path),
        Inches(box.left),
        Inches(box.top),
        Inches(box.width),
        Inches(box.height),
    )
    if mode == MODE_FILL:
        _apply_fill_crop(picture, region, ratio)
    elif crop_bottom:
        picture.crop_bottom = crop_bottom
    picture.name = f"{PIC_NAME_PREFIX}{target.screen_id}"
    picture._element.nvPicPr.cNvPr.set(
        "descr",
        f"{REGION_TAG}{region.left:.4f},{region.top:.4f},{region.width:.4f},{region.height:.4f}",
    )

    old = target.element
    old.addprevious(picture._element)
    old_rids = old.xpath(".//@r:embed")
    old.getparent().remove(old)
    # python-pptx's drop_rel() counts only r:id, not r:embed, so check usage here.
    # A re-added identical image is deduplicated onto the same rId and must survive.
    for rid in set(old_rids):
        if not slide._element.xpath(f'.//@r:embed[. = "{rid}"]'):
            slide.part.rels.pop(rid)
    return placement, _effective_ppi(size, region, placement)


def _inside(shape, region: Box) -> bool:
    if shape.left is None or shape.width is None:
        return False
    box, m = _shape_box(shape), OVERLAY_MARGIN_IN
    return (
        box.left >= region.left - m
        and box.top >= region.top - m
        and box.left + box.width <= region.left + region.width + m
        and box.top + box.height <= region.top + region.height + m
    )


def _is_annotation(shape) -> bool:
    """Explanation marks the planner drew over the screen, as opposed to drawn UI.

    - callout badge: small filled auto shape holding just a number (1, 2a, 1-1)
    - highlight frame: empty auto shape with no fill and a dashed outline
    - pointer: a line or connector with an arrowhead or a dashed stroke
    A group counts only when everything in it is an annotation.
    """
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return all(_is_annotation(child) for child in shape.shapes)
    sp_pr = shape._element.find(qn("p:spPr"))
    line = sp_pr.find(qn("a:ln")) if sp_pr is not None else None
    dash = line.find(qn("a:prstDash")) if line is not None else None
    dashed = dash is not None and dash.get("val", "solid") != "solid"
    if shape.shape_type == MSO_SHAPE_TYPE.LINE or shape._element.tag.endswith("}cxnSp"):
        arrow = line is not None and any(
            end is not None and end.get("type", "none") != "none"
            for end in (line.find(qn("a:headEnd")), line.find(qn("a:tailEnd")))
        )
        return arrow or dashed
    if shape.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
        return False
    text = shape.text_frame.text.strip() if shape.has_text_frame else ""
    no_fill = sp_pr is not None and sp_pr.find(qn("a:noFill")) is not None
    small = _emu_to_in(shape.width) <= BADGE_MAX_IN and _emu_to_in(shape.height) <= BADGE_MAX_IN
    if text:
        return small and not no_fill and bool(_BADGE_TEXT_RE.fullmatch(text))
    return no_fill and dashed


def _header_top(slide, region: Box) -> Optional[float]:
    """Top of an app header bar drawn just above the screen picture, if any.

    Some decks put the app's top bar (logo, menu) in the slide layout or as a
    shape right above the screenshot. A Figma export includes that bar, so in
    clean mode the screen box grows upward to cover it.
    """
    candidates = list(slide.shapes)
    for owner in (slide.slide_layout, slide.slide_layout.slide_master):
        candidates += [s for s in owner.shapes if not s.is_placeholder]
    best: Optional[float] = None
    for shape in candidates:
        if shape.left is None or shape.width is None or shape.name.startswith(PIC_NAME_PREFIX):
            continue
        box = _shape_box(shape)
        bottom = box.top + box.height
        overlap = min(box.left + box.width, region.left + region.width) - max(box.left, region.left)
        if (
            box.top < region.top - HEADER_GAP_IN
            and abs(bottom - region.top) <= HEADER_GAP_IN
            and overlap >= region.width * HEADER_MIN_SHARE
            and box.width <= region.width * (2 - HEADER_MIN_SHARE)
        ):
            best = box.top if best is None else min(best, box.top)
    return best


def handle_overlay(slide, region: Box, clean: bool) -> tuple[int, int]:
    """Count (and in clean mode remove) shapes drawn over the screen region.

    Returns (shapes on the screen, annotations among them). Inserted Figma
    pictures are never touched; annotations are always kept.
    """
    on_screen = annotations = 0
    for shape in list(slide.shapes):
        if shape.name.startswith(PIC_NAME_PREFIX) or not _inside(shape, region):
            continue
        on_screen += 1
        if _is_annotation(shape):
            annotations += 1
        elif clean:
            element = shape._element
            rids = set(element.xpath(".//@r:embed"))
            element.getparent().remove(element)
            for rid in rids:
                if not slide._element.xpath(f'.//@r:embed[. = "{rid}"]'):
                    slide.part.rels.pop(rid)
    return on_screen, annotations


def run_insert(
    prs, images: dict[str, ImageEntry], mode: str, report: RunReport, overlay: str = OVERLAY_KEEP
) -> None:
    targets = find_targets(prs, report)
    used: set[str] = set()
    for target in targets:
        numbered = _key(f"{target.screen_id}_{target.occurrence}")
        image_key = numbered if numbered in images else _key(target.screen_id)
        image = images.get(image_key)
        if image is None:
            if target.kind == "marker":
                report.missing_image.append(target)
            elif target.kind in ("table", "textbox"):
                report.kept_existing.append(target)
            continue
        slide = prs.slides[target.slide_no - 1]
        if overlay == OVERLAY_CLEAN and target.kind != "picture":
            header = _header_top(slide, target.box)
            if header is not None:
                added = target.box.top - header
                target.box = Box(target.box.left, header, target.box.width, target.box.height + added)
                report.header_added[target.slide_no] = added
        placement, ppi = place_image(slide, target, image, mode)
        on_screen, annotations = handle_overlay(slide, target.box, overlay == OVERLAY_CLEAN)
        if on_screen:
            report.overlay[target.slide_no] = (on_screen, annotations)
        if ppi < MIN_PPI:
            width, height = _image_size(image.path)
            report.warnings.append(
                f"슬라이드 {target.slide_no}: `{image.path.name}` 해상도가 낮습니다 "
                f"({width}×{height}px, 표시 크기 기준 {ppi:.0f}ppi — {MIN_PPI}ppi 이상 권장). "
                "피그마에서 배율 2x로 다시 export하세요."
            )
        report.inserted.append((target, image, placement))
        used.add(image_key)
    report.unused_images = [img for key, img in sorted(images.items()) if key not in used]
    _warn_small_screens(targets, [t for t, _, _ in report.inserted], report)


def _natural_key(text: str) -> list:
    return [int(part) if part.isdigit() else part.upper() for part in _NUMBER_SPLIT_RE.split(text)]


def _id_table(slide):
    for shape in slide.shapes:
        if getattr(shape, "has_table", False) and shape.has_table:
            for row in shape.table.rows:
                if any(_label(c.text) == TABLE_ID_LABEL for c in row.cells):
                    return shape.table
    return None


def _label(text: str) -> str:
    return re.sub(r"\s+", "", text).upper()


def _set_text(text_frame, value: str) -> None:
    """Replace the text while keeping the first run's formatting."""
    paragraphs = text_frame.paragraphs
    for extra in paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    runs = paragraphs[0].runs
    if runs:
        runs[0].text = value
        for run in runs[1:]:
            run._r.getparent().remove(run._r)
    else:
        paragraphs[0].text = value


def _in_footer(shape, slide_height: float) -> bool:
    return _emu_to_in(shape.top) >= slide_height * FOOTER_ZONE


def _page_number_frames(slide, page_no: int, slide_height: float) -> list:
    """Text frames that show this slide's page number (footer box, "페이지" table cell)."""
    wanted = str(page_no)
    frames = []
    for shape in slide.shapes:
        if getattr(shape, "has_table", False) and shape.has_table:
            for row in shape.table.rows:
                cells = list(row.cells)
                for idx, cell in enumerate(cells[:-1]):
                    if _label(cell.text) == PAGE_LABEL and cells[idx + 1].text.strip() == wanted:
                        frames.append(cells[idx + 1].text_frame)
        elif (
            getattr(shape, "has_text_frame", False)
            and shape.text_frame.text.strip() == wanted
            and _in_footer(shape, slide_height)
        ):
            frames.append(shape.text_frame)
    return frames


def _screen_box(prs) -> Optional[Box]:
    """Most common box of the existing screenshot across screen slides."""
    counts: dict[tuple, int] = {}
    height = _emu_to_in(prs.slide_height)
    for slide in prs.slides:
        picture = _largest_picture(slide) if _slide_screen_id(slide, height)[0] else None
        if picture is not None:
            box = _shape_box(picture)
            key = tuple(round(v, 2) for v in (box.left, box.top, box.width, box.height))
            counts[key] = counts.get(key, 0) + 1
    if not counts:
        return None
    return Box(*max(counts, key=counts.get))


def _is_skeleton(shape, page_no: int, slide_area: float, slide_height: float, screen: Box) -> bool:
    """Shapes every screen page shares, judged by position rather than by deck.

    Kept: tables outside the screen area (page header table, description table
    beside the screen), small pictures (logos), text in the header band (page
    title, screen ID, requirement ID), the footer page number and the empty
    content frame. Anything drawn on the screen area belongs to that screen.
    """
    if shape.left is None or shape.width is None:
        return False
    area = _emu_to_in(shape.width) * _emu_to_in(shape.height)
    if getattr(shape, "has_table", False) and shape.has_table:
        return not _inside(shape, screen)
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return area < MIN_SCREEN_AREA_SQIN and not _inside(shape, screen)
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP or not getattr(shape, "has_text_frame", False):
        return False
    text = shape.text_frame.text.strip()
    if text == str(page_no) and _in_footer(shape, slide_height):
        return True
    if text and _emu_to_in(shape.top) < slide_height * HEADER_ZONE and not _inside(shape, screen):
        return True
    return shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and not text and area >= slide_area * FRAME_MIN_SHARE


def _copy_rels(element, src_part, dst_part) -> None:
    for node in element.iter():
        for attr in (f"{{{_R_NS}}}embed", f"{{{_R_NS}}}link", f"{{{_R_NS}}}id"):
            rid = node.get(attr)
            if not rid or rid not in src_part.rels:
                continue
            rel = src_part.rels[rid]
            if rel.is_external:
                new_rid = dst_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            else:
                new_rid = dst_part.relate_to(rel.target_part, rel.reltype)
            node.set(attr, new_rid)


def _type_value(slide) -> str:
    table = _id_table(slide)
    if table is None:
        return ""
    for row in table.rows:
        cells = list(row.cells)
        for idx, cell in enumerate(cells[:-1]):
            if _label(cell.text) == TYPE_LABEL:
                return cells[idx + 1].text.strip()
    return ""


def _clone_source_index(prs, screen_indexes: list[int], position: int) -> int:
    """Screen slide nearest to ``position`` whose 유형 is "변경"; the neighbor if none.

    Pages marked 유지 use a gray frame in these decks, which a new screen must not inherit.
    """
    changed = [i for i in screen_indexes if _type_value(prs.slides[i]) == CLONE_TYPE_VALUE]
    pool = changed or screen_indexes
    return min(pool, key=lambda i: (abs(i - (position - 0.5)), i))


def _clone_screen_slide(prs, src, src_page: int, position: int, screen: Box):
    """Add a slide at 0-based ``position`` holding only the shared page skeleton of ``src``."""
    slide = prs.slides.add_slide(src.slide_layout)
    tree = slide.shapes._spTree
    for element in list(tree)[2:]:  # keep nvGrpSpPr / grpSpPr
        tree.remove(element)
    src_bg = src._element.cSld.bg
    if src_bg is not None:
        slide._element.cSld.insert(0, copy.deepcopy(src_bg))
    slide_height = _emu_to_in(prs.slide_height)
    slide_area = _emu_to_in(prs.slide_width) * slide_height
    for shape in src.shapes:
        if _is_skeleton(shape, src_page, slide_area, slide_height, screen):
            element = copy.deepcopy(shape._element)
            _copy_rels(element, src.part, slide.part)
            tree.append(element)
    id_list = prs.slides._sldIdLst
    entry = id_list[-1]
    id_list.remove(entry)
    id_list.insert(position, entry)
    return slide


def _fill_header_texts(slide, screen_id: str, slide_height: float) -> None:
    """Header-text decks: new ID in the ID box, placeholders for page-specific header text.

    Short tags (version "1.9", progress "9/12") are left as they are.
    """
    header = [
        shape for shape in slide.shapes
        if getattr(shape, "has_text_frame", False)
        and shape.text_frame.text.strip()
        and _emu_to_in(shape.top) < slide_height * HEADER_ZONE
    ]
    ids = [s for s in header if _ID_TEXT_RE.fullmatch(s.text_frame.text.strip())]
    id_box = min(ids, key=lambda s: s.left) if ids else None
    for shape in header:
        if shape is id_box:
            _set_text(shape.text_frame, screen_id)
        elif len(shape.text_frame.text.strip()) > HEADER_TAG_MAX_CHARS:
            _set_text(shape.text_frame, PLACEHOLDER)


def _reset_side_tables(slide, screen: Box) -> bool:
    """Description tables right of the screen: keep one row, first cell "1", rest placeholder."""
    found = False
    for shape in slide.shapes:
        if not (getattr(shape, "has_table", False) and shape.has_table):
            continue
        if _emu_to_in(shape.left) < screen.left + screen.width - OVERLAY_MARGIN_IN:
            continue
        found = True
        for extra in shape.table._tbl.tr_lst[1:]:
            extra.getparent().remove(extra)
        cells = list(shape.table.rows[0].cells)
        for idx, cell in enumerate(cells):
            _set_text(cell.text_frame, "1" if idx == 0 and len(cells) > 1 else PLACEHOLDER)
        shape.height = shape.table.rows[0].height
    return found


def _fill_new_slide(slide, screen_id: str, page_no: int, src_page: int, screen: Box, prs) -> None:
    slide_height = _emu_to_in(prs.slide_height)
    table = _id_table(slide)
    if table is not None:
        for row in table.rows:
            cells = list(row.cells)
            for idx, cell in enumerate(cells[:-1]):
                label = _label(cell.text)
                if label == TABLE_ID_LABEL:
                    _set_text(cells[idx + 1].text_frame, screen_id)
                elif label in UNKNOWN_FIELD_LABELS:
                    _set_text(cells[idx + 1].text_frame, PLACEHOLDER)
    else:
        _fill_header_texts(slide, screen_id, slide_height)
    for frame in _page_number_frames(slide, src_page, slide_height):
        _set_text(frame, str(page_no))

    if _reset_side_tables(slide, screen):
        _add_marker(slide, screen, screen_id)
        return

    frame_right = _emu_to_in(prs.slide_width) - 0.3
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            frame_right = _emu_to_in(shape.left + shape.width) - 0.1
    left = screen.left + screen.width + 0.3
    note = slide.shapes.add_textbox(
        Inches(left), Inches(screen.top), Inches(max(frame_right - left, 1.0)), Inches(screen.height)
    )
    note.name = "FIGMA-NOTE"
    note.text_frame.word_wrap = True
    for idx, line in enumerate(DESCRIPTION_PLACEHOLDER):
        paragraph = note.text_frame.paragraphs[0] if idx == 0 else note.text_frame.add_paragraph()
        paragraph.text = line
        for run in paragraph.runs:
            run.font.name = FONT_FAMILY
            run.font.size = Pt(10)

    _add_marker(slide, screen, screen_id)


def _add_marker(slide, screen: Box, screen_id: str) -> None:
    marker = slide.shapes.add_textbox(
        Inches(screen.left), Inches(screen.top), Inches(screen.width), Inches(screen.height)
    )
    marker.text_frame.text = f"[화면 교체 위치: {screen_id}]"


def add_missing_slides(prs, images: dict[str, ImageEntry], report: RunReport) -> list[str]:
    """Create a screen page for each image whose ID has no slide yet.

    The new page clones the shared skeleton of the neighboring screen slide,
    placed in natural ID order, and carries a marker for the regular insert
    pass. Page numbers of later slides are shifted to match.
    """
    present = {_key(t.screen_id) for t in find_targets(prs, RunReport())}

    def is_new(key: str) -> bool:
        numbered = _OCCURRENCE_RE.match(key)
        return key not in present and not (numbered and numbered.group("base") in present)

    new_ids = sorted((img.screen_id for key, img in images.items() if is_new(key)), key=_natural_key)
    if not new_ids:
        return []
    screen = _screen_box(prs)
    if screen is None:
        report.warnings.append(
            "화면 ID가 있는 화면 슬라이드가 없어 새 슬라이드를 만들지 않았습니다 "
            f"({', '.join(new_ids)})."
        )
        return []

    slide_height = _emu_to_in(prs.slide_height)
    numbered = {
        slide.slide_id: _page_number_frames(slide, page, slide_height)
        for page, slide in enumerate(prs.slides, start=1)
    }
    added: list[str] = []
    for screen_id in new_ids:
        ids = [(i, _slide_screen_id(s, slide_height)[0]) for i, s in enumerate(prs.slides)]
        screens = [(i, sid) for i, sid in ids if sid]
        before = [i for i, sid in screens if _natural_key(sid) < _natural_key(screen_id)]
        neighbor = before[-1] if before else screens[0][0]
        position = neighbor + 1 if before else neighbor
        src_index = _clone_source_index(prs, [i for i, _ in screens], position)
        src = prs.slides[src_index]
        src_page = src_index + 1  # page numbers are checked against the slide position
        slide = _clone_screen_slide(prs, src, src_page, position, screen)
        _fill_new_slide(slide, screen_id, position + 1, src_page, screen, prs)
        added.append(screen_id)

    for page, slide in enumerate(prs.slides, start=1):
        for frame in numbered.get(slide.slide_id, []):
            _set_text(frame, str(page))
    report.warnings.append(
        "새 슬라이드를 추가했습니다 — 목차·개정 이력 등 다른 페이지의 화면 목록은 자동으로 바뀌지 않으니 확인하세요."
    )
    return added


def find_template(template_dir: Path) -> Path:
    candidates = [
        p for p in sorted(template_dir.glob("*.pptx")) if not p.name.startswith("~$")
    ]
    if len(candidates) != 1:
        found = ", ".join(p.name for p in candidates) or "없음"
        raise RuntimeError(
            f"`{TEMPLATE_DIR}` 폴더에는 .pptx 파일이 정확히 1개 있어야 합니다 (현재: {found})."
        )
    return candidates[0]


def _unique_path(path: Path) -> Path:
    candidate, n = path, 2
    while candidate.exists():
        candidate = path.with_name(f"{path.stem}_{n}{path.suffix}")
        n += 1
    return candidate


def save_with_archive(prs, template: Path, output_dir: Path) -> tuple[Path, Optional[Path], list[Path]]:
    """Save the new final deck; move the previous one to _archive and prune old archives."""
    output_dir.mkdir(parents=True, exist_ok=True)
    final_path = output_dir / f"{template.stem}{OUTPUT_SUFFIX}.pptx"
    tmp_path = output_dir / f".{template.stem}{OUTPUT_SUFFIX}.tmp.pptx"
    prs.save(str(tmp_path))

    archived: Optional[Path] = None
    pruned: list[Path] = []
    if final_path.exists():
        archive_dir = output_dir / ARCHIVE_DIR
        archive_dir.mkdir(exist_ok=True)
        stamp = datetime.fromtimestamp(final_path.stat().st_mtime).strftime("%Y%m%d_%H%M")
        archived = _unique_path(archive_dir / f"{template.stem}_{stamp}.pptx")
        # Same-volume rename: fails cleanly if PowerPoint holds the file, unlike
        # shutil.move's copy+delete fallback that leaves a stray archive copy.
        try:
            final_path.rename(archived)
        except PermissionError:
            tmp_path.unlink(missing_ok=True)
            raise
        old = sorted(archive_dir.glob("*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)
        for path in old[ARCHIVE_KEEP:]:
            path.unlink()
            pruned.append(path)
    tmp_path.replace(final_path)
    return final_path, archived, pruned


def _rel(path: Path, root: Path) -> str:
    try:
        return path.relative_to(root).as_posix()
    except ValueError:
        return str(path)


def format_report(
    report: RunReport,
    *,
    added: list[str],
    root: Path,
    template: Path,
    mode: str,
    started: datetime,
    final_path: Optional[Path],
    archived: Optional[Path],
    pruned: list[Path],
) -> str:
    report_overlay_mode = report.overlay_mode
    mode_label = "비율 유지 맞춤(fit)" if mode == MODE_FIT else "영역 채우기·잘림(fill)"
    lines = [
        f"## {started:%Y-%m-%d %H:%M} 실행",
        "",
        f"- 프로젝트: `{root.name}`",
        f"- 템플릿: `{_rel(template, root)}`",
        f"- 배치 방식: {mode_label}",
        f"- 결과 파일: `{_rel(final_path, root)}`" if final_path else "- 결과 파일: (미리보기 실행 — 저장 안 함)",
    ]
    if archived:
        lines.append(f"- 이전 최종본 보관: `{_rel(archived, root)}`")
    for path in pruned:
        lines.append(f"- 보관 개수 초과로 삭제: `{_rel(path, root)}`")
    imported_title = "다운로드에서 가져옴" if final_path else "다운로드에서 가져올 예정 (미리보기 — 옮기지 않음)"
    lines += ["", f"### {imported_title} ({len(report.imported)})", ""]
    lines += [f"- `{src}` → `{dest}`" for src, dest in report.imported] or ["- 없음"]
    if report.recent_downloads:
        lines += ["", f"### 다운로드 — 최근 파일, 가져오지 않음 ({len(report.recent_downloads)})", ""]
        lines += [f"- `{name}` ({day:%m-%d})" for name, day in report.recent_downloads]
        lines.append("- 넣으려면 `--download-days 3`으로 다시 실행하세요.")
    if report.moved:
        lines += ["", f"### 이미지 폴더 정리 ({len(report.moved)})", ""]
        lines += [f"- `{old}` → `{IMAGES_DIR}/{new}`" for old, new in report.moved]
    if report.archived:
        lines += ["", f"### 이전 이미지 보관 ({len(report.archived)})", ""]
        lines += [f"- `{old}` → `{IMAGES_DIR}/{IMAGE_ARCHIVE_DIR}/{new}`" for old, new in report.archived]
    lines += ["", f"### 삽입 성공 ({len(report.inserted)})", ""]
    if report.inserted:
        lines += ["| 슬라이드 | 화면 ID | 사용 이미지 | 대상 | 배치 |", "|---|---|---|---|---|"]
        for target, image, placement in report.inserted:
            kind = {"marker": "마커", "picture": "삽입 이미지 교체", "table": "화면ID 표·기존 화면 교체",
                    "textbox": "화면ID 텍스트·기존 화면 교체"}[
                target.kind
            ]
            lines.append(
                f"| {target.slide_no} | {target.screen_id} | `{image.label}` | {kind} | {placement} |"
            )
    else:
        lines.append("- 없음")
    if report.overlay:
        clean = report_overlay_mode == OVERLAY_CLEAN
        title = "화면 위 도형 — 삭제(번호 배지·점선 테두리·화살표는 남김)" if clean else "화면 위 도형 — 그대로 남김"
        lines += ["", f"### {title} ({len(report.overlay)})", ""]
        for slide_no, (on_screen, kept) in sorted(report.overlay.items()):
            others = on_screen - kept
            if clean:
                header = report.header_added.get(slide_no)
                note = f", 위쪽 헤더 {header:.2f}in까지 화면 영역에 포함" if header else ""
                lines.append(f"- 슬라이드 {slide_no}: {others}개 삭제, 번호 배지 등 {kept}개 유지{note}")
            else:
                lines.append(f"- 슬라이드 {slide_no}: {on_screen}개 남음 (번호 배지 등 {kept}개, 그 외 {others}개)")
    if added:
        new_pages = [t for t, _, _ in report.inserted if _key(t.screen_id) in {_key(a) for a in added}]
        lines += ["", f"### 새 슬라이드 추가 ({len(added)})", ""]
        lines += [
            f"- 슬라이드 {t.slide_no}: `{t.screen_id}` — `{PLACEHOLDER}`로 표시된 항목(화면명·경로·설명 등)을 채워 주세요"
            for t in new_pages
        ]
    lines += ["", f"### 매칭 실패 — 마커는 있는데 이미지 없음 ({len(report.missing_image)})", ""]
    lines += [f"- 슬라이드 {t.slide_no}: `{t.screen_id}`" for t in report.missing_image] or ["- 없음"]
    lines += ["", f"### 이미지 없음 — 화면ID 슬라이드, 기존 화면 유지 ({len(report.kept_existing)})", ""]
    lines += [f"- 슬라이드 {t.slide_no}: `{t.screen_id}`" for t in report.kept_existing] or ["- 없음"]
    lines += ["", f"### 매칭 실패 — 이미지는 있는데 마커·화면ID 없음 ({len(report.unused_images)})", ""]
    lines += [f"- `{IMAGES_DIR}/{i.label}`" for i in report.unused_images] or ["- 없음"]
    if report.warnings:
        lines += ["", f"### 확인 필요 ({len(report.warnings)})", ""]
        lines += [f"- {w}" for w in report.warnings]
    return "\n".join(lines) + "\n"


def append_log(log_path: Path, entry: str) -> None:
    if not log_path.exists():
        log_path.write_text(
            "# 피그마 화면 삽입 실행 이력\n\n최근 실행이 맨 아래에 추가됩니다.\n",
            encoding="utf-8",
        )
    with log_path.open("a", encoding="utf-8", newline="\n") as fh:
        fh.write("\n" + entry)


def configure_utf8_stdio() -> None:
    """Force UTF-8 console output so Korean status text survives non-UTF-8 Windows locales."""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def list_projects() -> list[str]:
    if not PROJECTS_DIR.is_dir():
        return []
    return sorted(p.name for p in PROJECTS_DIR.iterdir() if p.is_dir() and not p.name.startswith("."))


def create_project(name: str) -> Path:
    """Create the folder skeleton for a new project and return its root."""
    clean = name.strip()
    if not clean or clean.startswith(".") or clean.endswith(".") or _BAD_PROJECT_CHARS_RE.search(clean):
        raise RuntimeError(
            f"프로젝트 이름 `{name}`은 쓸 수 없습니다. "
            '\\ / : * ? " < > | 문자와 앞뒤 점(.)은 빼 주세요.'
        )
    root = PROJECTS_DIR / clean
    if root.exists():
        raise RuntimeError(f"`projects/{clean}` 프로젝트가 이미 있습니다.")
    for sub in (TEMPLATE_DIR, IMAGES_DIR, OUTPUT_DIR):
        (root / sub).mkdir(parents=True)
    return root


def resolve_project(name: Optional[str]) -> Path:
    """Return the project root for --project, or the only project when omitted."""
    projects = list_projects()
    available = ", ".join(projects) or "없음"
    if name:
        root = PROJECTS_DIR / name.strip()
        if not root.is_dir():
            raise RuntimeError(
                f"`projects/{name}` 프로젝트가 없습니다 (현재 프로젝트: {available}). "
                "새로 만들려면 --new-project를 쓰세요."
            )
        return root
    if len(projects) == 1:
        return PROJECTS_DIR / projects[0]
    if not projects:
        raise RuntimeError("프로젝트가 없습니다. --new-project <이름>으로 먼저 만드세요.")
    raise RuntimeError(f"프로젝트를 --project로 지정하세요 (현재 프로젝트: {available}).")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Insert Figma-exported screens into PPT marker shapes.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    target = parser.add_mutually_exclusive_group()
    target.add_argument(
        "--project",
        help="project folder name under projects/ (optional when only one exists)",
    )
    target.add_argument(
        "--new-project",
        metavar="NAME",
        help="create projects/NAME with empty 01_template / 02_images / 03_output and exit",
    )
    target.add_argument("--list", action="store_true", help="list projects and exit")
    target.add_argument(
        "--root",
        type=Path,
        help="use this folder directly as the project root (for testing)",
    )
    parser.add_argument(
        "--mode",
        choices=(MODE_FIT, MODE_FILL),
        default=MODE_FIT,
        help="fit: keep ratio inside the box (default); fill: cover the box and crop",
    )
    parser.add_argument(
        "--overlay",
        choices=(OVERLAY_KEEP, OVERLAY_CLEAN),
        default=OVERLAY_KEEP,
        help="shapes drawn over a replaced screen: keep (default) or clean "
        "(remove them, keeping callout badges, dashed frames and arrows)",
    )
    parser.add_argument(
        "--add-slides",
        action="store_true",
        help="(default, kept for compatibility) add a screen page for new IDs",
    )
    parser.add_argument(
        "--no-add-slides",
        action="store_true",
        help="do not add pages for new IDs; report those images as unmatched",
    )
    parser.add_argument(
        "--downloads",
        type=Path,
        default=DOWNLOADS_DIR,
        help=f"folder to pick up Figma exports from (default: {DOWNLOADS_DIR})",
    )
    parser.add_argument(
        "--download-days",
        type=int,
        default=DOWNLOAD_DAYS,
        metavar="N",
        help="import exports downloaded within the last N days (default: 1 = today)",
    )
    parser.add_argument(
        "--no-downloads",
        action="store_true",
        help="do not import from the downloads folder; use 02_images only",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="report matches only; do not save, archive, or log",
    )
    return parser


def main(argv: Optional[list[str]] = None) -> int:
    configure_utf8_stdio()
    args = build_parser().parse_args(argv)
    if args.list:
        print("\n".join(list_projects()) or "(프로젝트 없음)")
        return 0
    try:
        if args.new_project:
            root = create_project(args.new_project)
            print(
                f"`projects/{root.name}` 프로젝트를 만들었습니다.\n"
                f"- 원본 PPT 1개 → projects/{root.name}/{TEMPLATE_DIR}/\n"
                f"- 피그마 이미지 → projects/{root.name}/{IMAGES_DIR}/ (다운로드에서 자동으로 가져옴)"
            )
            return 0
        root = args.root.resolve() if args.root else resolve_project(args.project)
        template = find_template(root / TEMPLATE_DIR)
    except RuntimeError as exc:
        print(f"오류: {exc}", file=sys.stderr)
        return 1
    started = datetime.now()

    report = RunReport(overlay_mode=args.overlay)
    prs = Presentation(str(template))
    downloads = (
        []
        if args.no_downloads
        else _scan_downloads(
            args.downloads, _IdMatcher(deck_screen_ids(prs)), report, max(args.download_days, 1)
        )
    )
    with tempfile.TemporaryDirectory() as scratch:  # dry-run copies of zip members
        images = organize_images(
            root / IMAGES_DIR, downloads, report, dry_run=args.dry_run, scratch=Path(scratch)
        )
        added = [] if args.no_add_slides else add_missing_slides(prs, images, report)
        run_insert(prs, images, args.mode, report, args.overlay)

    final_path: Optional[Path] = None
    archived: Optional[Path] = None
    pruned: list[Path] = []
    if not args.dry_run:
        try:
            final_path, archived, pruned = save_with_archive(prs, template, root / OUTPUT_DIR)
        except PermissionError as exc:
            print(
                f"오류: 파일을 쓸 수 없습니다 ({exc.filename}). "
                "PowerPoint에서 결과 파일을 닫고 다시 실행하세요.",
                file=sys.stderr,
            )
            return 1

    entry = format_report(
        report,
        added=added,
        root=root,
        template=template,
        mode=args.mode,
        started=started,
        final_path=final_path,
        archived=archived,
        pruned=pruned,
    )
    if not args.dry_run:
        append_log(root / LOG_FILE, entry)
    print(entry)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
