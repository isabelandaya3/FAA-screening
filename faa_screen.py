"""Screen structures on the current FAA OE/AAA Pre-Screening Tool."""

from __future__ import annotations

import re
import time
from pathlib import Path

from dataclasses import dataclass
from datetime import datetime
from hashlib import sha256
from math import ceil

from openpyxl import load_workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeout
from playwright.sync_api import sync_playwright
from PIL import Image

# ---------------------------------------------------------------------------
# Edit these paths before you run the script.
# ---------------------------------------------------------------------------
WORKBOOK_PATH = r"C:\Users\HP\Downloads\JOB NO - LINE NAME FAA Screening.xlsm"
PDF_FOLDER = r"C:\Users\HP\OneDrive - scu.edu\Documents\FAA"
# ---------------------------------------------------------------------------

DATUM = "NAD83"
SHEET = "FAA Criteria Tool"
DATA_SHEET = "Data Sheet"
MAX_FILE_ROWS = 250
PRESCREEN_URL = "https://oeaaa.faa.gov/oeaaa/oe3a/main/#/noticePrescreen"
STRUCTURE_TYPE_ID = "TRANSMISSION_LINE$T_L_TOWER"
RESULT_RE = re.compile(
    r"(required to file|not required to file|does not exceed|exceed)",
    re.I,
)


@dataclass
class Structure:
    row: int
    number: str
    pls_file: str
    lon_deg: float | None
    lat_deg: float | None
    elevation_ft: float
    height_ft: float
    lon_dms: str
    lat_dms: str
    traverseway: str
    on_airport: str
    structure_type: str = "TRANSMISSION_LINE$T_L_TOWER"
    last_result: str = ""
    last_hash: str = ""

    @property
    def elevation_whole(self) -> int:
        return int(round_half_up(self.elevation_ft))

    @property
    def height_whole(self) -> int:
        if float(self.height_ft).is_integer():
            return int(self.height_ft)
        return int(ceil(self.height_ft))


def round_half_up(value: float) -> int:
    return int(value + 0.5) if value >= 0 else int(value - 0.5)


def _as_float(value) -> float | None:
    if value is None or value == "":
        return None
    return float(value)


def parse_dms(text: str) -> tuple[str, str]:
    raw = str(text).strip()
    hemi = raw[-1].upper() if raw[-1:] in "NSEWnsew" else ""
    body = raw[:-1] if hemi else raw
    match = re.search(
        r"(-?\d+(?:\.\d+)?)\s*[dD°]\s*(\d+(?:\.\d+)?)\s*['′]\s*(\d+(?:\.\d+)?)",
        body,
    )
    if match:
        deg = int(float(match.group(1)))
        minutes = int(float(match.group(2)))
        seconds = round(float(match.group(3)), 2)
        return f"{abs(deg)}-{minutes:02d}-{seconds:05.2f}", hemi
    hyphen = re.match(
        r"(-?\d+)\s*-\s*(\d+)\s*-\s*(\d+(?:\.\d+)?)\s*([NSEW])?$",
        raw,
        re.I,
    )
    if hyphen:
        deg = int(hyphen.group(1))
        minutes = int(hyphen.group(2))
        seconds = float(hyphen.group(3))
        hemi = (hyphen.group(4) or hemi).upper()
        return f"{abs(deg)}-{minutes:02d}-{seconds:05.2f}", hemi
    raise ValueError(f"Unrecognized DMS value: {text!r}")


def load_structures(workbook_path: Path) -> list[Structure]:
    wb = load_workbook(workbook_path, data_only=True)
    ws = wb[SHEET]
    rows: list[Structure] = []
    for r in range(2, ws.max_row + 1):
        number = ws[f"B{r}"].value
        if number is None or str(number).strip() == "":
            continue
        rows.append(
            Structure(
                row=r,
                number=str(number).strip(),
                pls_file=str(ws[f"C{r}"].value or "").strip(),
                lon_deg=_as_float(ws[f"D{r}"].value),
                lat_deg=_as_float(ws[f"E{r}"].value),
                elevation_ft=float(ws[f"F{r}"].value),
                height_ft=float(ws[f"G{r}"].value),
                lon_dms=str(ws[f"H{r}"].value).strip(),
                lat_dms=str(ws[f"I{r}"].value).strip(),
                traverseway=str(ws[f"J{r}"].value or "No Traverseway").strip(),
                on_airport=str(ws[f"K{r}"].value or "No").strip(),
                last_result=str(ws[f"L{r}"].value or "").strip(),
                last_hash=str(ws[f"N{r}"].value or "").strip(),
            )
        )
    return rows


def input_fingerprint(structure: Structure, datum: str) -> str:
    payload = "|".join(
        [
            structure.number,
            str(structure.lat_dms),
            str(structure.lon_dms),
            str(structure.lat_deg),
            str(structure.lon_deg),
            str(structure.elevation_ft),
            str(structure.height_ft),
            str(structure.on_airport),
            str(structure.structure_type),
            datum,
        ]
    )
    return sha256(payload.encode("utf-8")).hexdigest()[:16]


def write_result(
    workbook_path: Path,
    structure: Structure,
    result_text: str,
    requires_filing: bool,
) -> None:
    wb = load_workbook(workbook_path, keep_vba=workbook_path.suffix.lower() == ".xlsm")
    ws = wb[SHEET]
    ws[f"L{structure.row}"] = result_text
    ws[f"M{structure.row}"] = datetime.now()
    ws[f"N{structure.row}"] = input_fingerprint(structure, DATUM)

    data = wb[DATA_SHEET]
    if requires_filing:
        lat_fmt, _ = parse_dms(structure.lat_dms)
        lon_fmt, _ = parse_dms(structure.lon_dms)
        next_row = 2
        while next_row <= MAX_FILE_ROWS + 1 and data[f"A{next_row}"].value:
            if str(data[f"A{next_row}"].value).strip() == structure.number:
                break
            next_row += 1
        if next_row > MAX_FILE_ROWS + 1:
            raise RuntimeError("Data Sheet already has 250 cases.")
        data[f"A{next_row}"] = structure.number
        data[f"B{next_row}"] = lat_fmt
        data[f"C{next_row}"] = lon_fmt
        data[f"D{next_row}"] = structure.elevation_whole
        data[f"E{next_row}"] = structure.height_whole
    wb.save(workbook_path)


def safe_structure_name(number: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "-", str(number).strip())
    return cleaned.strip(" .-") or "structure"


def pdf_path(folder: Path, number: str, requires_filing: bool) -> Path:
    suffix = "NEED TO FILE" if requires_filing else "FILE NOT REQUIRED"
    return folder / f"{safe_structure_name(number)}_FAA Screening_{suffix}.pdf"


def remove_stale_pdfs(folder: Path, number: str, keep: Path | None = None) -> None:
    prefix = f"{safe_structure_name(number)}_FAA Screening_"
    for path in folder.glob(f"{prefix}*.pdf"):
        if keep is None or path.resolve() != keep.resolve():
            path.unlink(missing_ok=True)


def should_screen(structure: Structure, folder: Path) -> bool:
    digest = input_fingerprint(structure, DATUM)
    if structure.last_hash != digest or not structure.last_result:
        return True
    need = pdf_path(folder, structure.number, True)
    skip = pdf_path(folder, structure.number, False)
    return not (need.exists() or skip.exists())


def dismiss_overlays(page) -> None:
    for name in ("Stay Logged In", "Continue", "OK", "Close"):
        btn = page.get_by_role("button", name=name)
        try:
            if btn.count() and btn.first.is_visible():
                btn.first.click(timeout=1500)
                time.sleep(0.3)
        except Exception:
            pass


def hide_modals(page) -> None:
    page.evaluate(
        """() => {
            if (typeof $ !== 'undefined') {
                $('#verifyNoticePointModal').modal('hide');
                $('#confirmPointModal').modal('hide');
            }
        }"""
    )


def select_structure_type(page) -> None:
    page.locator("structure-type").click()
    page.wait_for_timeout(500)
    leaf = page.locator(f'span[data-id="{STRUCTURE_TYPE_ID}"]')
    if leaf.count() == 0:
        page.get_by_text("T-L Tower", exact=False).first.click()
        return
    leaf.first.scroll_into_view_if_needed()
    leaf.first.click()
    page.wait_for_timeout(400)


def apply_user_elevation_and_datum(page, structure: Structure) -> None:
    page.evaluate(
        """([elev, datum]) => {
            const el = document.querySelector('notice-criteria');
            const scope = angular.element(el).isolateScope() || angular.element(el).scope();
            const point = scope.data.points.find(p => p.include) || scope.data.points[0];
            point.datum = datum;
            point.siteElevation = Number(elev);
            point.siteElevationSource = 'USER';
            point.elevationDetails = {
                validation: 'PASSED',
                comments: 'Project spreadsheet / PLS-CADD survey elevation'
            };
            if (scope.formData) scope.formData.datum = datum;
            scope.$apply();
        }""",
        [structure.elevation_whole, DATUM],
    )


def fill_point(page, structure: Structure) -> None:
    lat_body, lat_dir = parse_dms(structure.lat_dms)
    lon_body, lon_dir = parse_dms(structure.lon_dms)
    if not lat_dir:
        lat_dir = "N"
    if not lon_dir:
        lon_dir = "W"

    page.evaluate(
        """([lat, latDir, lon, lonDir, height, elev, onAirport, datum]) => {
            const el = document.querySelector('notice-criteria');
            const scope = angular.element(el).isolateScope() || angular.element(el).scope();
            const point = scope.data.points.find(p => p.include) || scope.data.points[0];
            point.formLat2 = lat;
            point.latDir = latDir;
            point.formLon2 = lon;
            point.lonDir = lonDir;
            point.datum = datum;
            point.structureHeight = Number(height);
            point.siteElevation = Number(elev);
            point.siteElevationSource = 'USER';
            point.elevationDetails = {
                validation: 'PASSED',
                comments: 'Project spreadsheet / PLS-CADD survey elevation'
            };
            point.state = 'COORD_DIRTY';
            scope.data.onAirport = !!onAirport;
            if (scope.formData) scope.formData.datum = datum;
            scope.$apply();
            if (scope.redrawPoint) scope.redrawPoint(point);
            return true;
        }""",
        [
            lat_body,
            lat_dir,
            lon_body,
            lon_dir,
            structure.height_whole,
            structure.elevation_whole,
            structure.on_airport.lower() == "yes",
            DATUM,
        ],
    )
    page.wait_for_timeout(2000)
    validate = page.get_by_role("button", name="VALIDATE")
    if validate.count() and validate.first.is_visible():
        validate.first.click()
        page.wait_for_timeout(1200)
    handle_point_modal(page, structure)
    apply_user_elevation_and_datum(page, structure)
    if structure.lat_deg is not None and structure.lon_deg is not None:
        page.evaluate(
            """([lat, lon, elev, datum]) => {
                const el = document.querySelector('notice-criteria');
                const scope = angular.element(el).isolateScope() || angular.element(el).scope();
                const point = scope.data.points.find(p => p.include) || scope.data.points[0];
                if (!point.lat) point.lat = lat;
                if (!point.lon) point.lon = lon;
                point.datum = datum;
                point.siteElevation = Number(elev);
                point.siteElevationSource = 'USER';
                if (point.elevationDetails) {
                    point.elevationDetails.validation = 'PASSED';
                    point.elevationDetails.comments = 'Project spreadsheet / PLS-CADD survey elevation';
                }
                if (scope.isPointValid) scope.isPointValid(point);
                scope.$apply();
            }""",
            [structure.lat_deg, structure.lon_deg, structure.elevation_whole, DATUM],
        )
    page.wait_for_timeout(400)


def handle_point_modal(page, structure: Structure) -> None:
    modal = page.locator("#verifyNoticePointModal")
    try:
        modal.wait_for(state="visible", timeout=4000)
    except PlaywrightTimeout:
        return
    own = page.locator("#customRadioInline2")
    if own.count():
        own.check()
        elev_box = page.locator(
            "input[ng-model='noticeData.verifyPointModalPoint.userSiteElev']"
        )
        elev_box.fill(str(structure.elevation_whole))
        comment = page.locator("#verifyNoticePointModal textarea")
        if comment.count():
            comment.fill("Project spreadsheet / PLS-CADD survey elevation")
        page.evaluate(
            """([elev, datum]) => {
                const el = document.querySelector('notice-criteria');
                const scope = angular.element(el).isolateScope() || angular.element(el).scope();
                if (scope.noticeData && scope.noticeData.verifyPointModalPoint) {
                    scope.noticeData.verifyPointModalPoint.siteElevationSource = 'USER';
                    scope.noticeData.verifyPointModalPoint.userSiteElev = Number(elev);
                    scope.noticeData.verifyPointModalPoint.datum = datum;
                }
                scope.$apply();
            }""",
            [structure.elevation_whole, DATUM],
        )
    page.get_by_role("button", name="Accept Point").click()
    page.wait_for_timeout(800)
    apply_user_elevation_and_datum(page, structure)


def read_scope_results(page):
    return page.evaluate(
        """() => {
            const el = document.querySelector('notice-criteria');
            if (!el || typeof angular === 'undefined') return null;
            const scope = angular.element(el).isolateScope() || angular.element(el).scope();
            if (!scope) return null;
            return {
                exceed: scope.preScreenResults ? scope.preScreenResults.isExceedence : null,
                results: scope.results || [],
                header: scope.resultHeader || ''
            };
        }"""
    )


def compose_result_text(payload: dict) -> str:
    exceed = payload.get("exceed") or []
    messages = [str(m).strip() for m in (payload.get("results") or []) if str(m).strip()]
    header = str(payload.get("header") or "").strip()
    if any(bool(x) for x in exceed):
        lead = (
            header
            or "Based on the information you provided, you are required to file notice with the FAA."
        )
    elif isinstance(exceed, list) and len(exceed) > 0:
        lead = (
            header
            or "Based on the information you provided, you are not required to file notice with the FAA."
        )
    else:
        lead = header or ""
    if not lead and messages:
        lead = " ".join(messages)
    parts = [lead] + [m for m in messages if m not in lead]
    return "\n".join(p for p in parts if p)


def submit_and_read(page) -> str:
    invoked = page.evaluate(
        """() => {
            const el = document.querySelector('notice-criteria');
            const scope = angular.element(el).isolateScope() || angular.element(el).scope();
            if (!scope || !scope.noticeCriteria) return false;
            scope.noticeCriteria();
            scope.$applyAsync();
            return true;
        }"""
    )
    if not invoked:
        page.get_by_role("button", name="Submit").click()

    deadline = time.time() + 60
    payload = None
    while time.time() < deadline:
        payload = read_scope_results(page)
        exceed = payload.get("exceed") if payload else None
        if isinstance(exceed, list) and len(exceed) > 0:
            break
        body = page.locator("body").inner_text()
        if RESULT_RE.search(body) and "Pre-Screening Tool" in body:
            if "required to file" in body.lower() or "does not exceed" in body.lower():
                break
        page.wait_for_timeout(500)
    else:
        raise RuntimeError(f"FAA pre-screen did not return results: {payload!r}")

    text = compose_result_text(payload or {})
    if not RESULT_RE.search(text):
        visible = page.locator("body").inner_text()
        matches = [ln.strip() for ln in visible.splitlines() if RESULT_RE.search(ln)]
        text = "\n".join(matches) or text
    if not RESULT_RE.search(text):
        raise RuntimeError(f"No FAA result text. scope={payload!r}")
    return " ".join(text.split())


def requires_filing(result_text: str) -> bool:
    lower = result_text.lower()
    if "not required to file" in lower or "does not exceed" in lower:
        return False
    if "required to file" in lower or "exceed" in lower:
        return True
    return False


def save_result_pdf(page, dest: Path) -> None:
    dest.parent.mkdir(parents=True, exist_ok=True)
    hide_modals(page)
    page.wait_for_timeout(400)
    png_path = dest.with_suffix(".png")
    page.screenshot(path=str(png_path), full_page=True)
    image = Image.open(png_path).convert("RGB")
    image.save(dest, "PDF", resolution=150)
    png_path.unlink(missing_ok=True)
    if not dest.exists() or dest.stat().st_size < 1000:
        raise RuntimeError(f"PDF was not written: {dest}")


def screen_one(page, structure: Structure, pdf_folder: Path) -> tuple[str, Path]:
    try:
        print("  opening pre-screen ...")
        page.goto(PRESCREEN_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(3500)
        dismiss_overlays(page)
        page.wait_for_selector("structure-type", timeout=30000)
        print("  selecting structure type ...")
        select_structure_type(page)
        print(f"  filling NAD83 point with spreadsheet elevation {structure.elevation_whole} ft ...")
        fill_point(page, structure)
        print("  submitting ...")
        result = submit_and_read(page)
        if not RESULT_RE.search(result):
            raise RuntimeError(f"No FAA result text for {structure.number}: {result[:300]!r}")
        apply_user_elevation_and_datum(page, structure)
        hide_modals(page)
        filing = requires_filing(result)
        dest = pdf_path(pdf_folder, structure.number, filing)
        remove_stale_pdfs(pdf_folder, structure.number, keep=None)
        save_result_pdf(page, dest)
        print(f"  saved {dest.name}")
        return result, dest
    except Exception as exc:
        print(f"  ERROR {structure.number}: {type(exc).__name__}: {exc}")
        try:
            if not page.is_closed():
                page.screenshot(
                    path=str(pdf_folder / f"{safe_structure_name(structure.number)}_FAILED.png"),
                    full_page=True,
                )
        except Exception as shot_exc:
            print(f"  Could not save failure screenshot: {shot_exc}")
        raise


def run(workbook: Path, pdf_folder: Path) -> list[tuple[Structure, str, bool]]:
    pdf_folder.mkdir(parents=True, exist_ok=True)
    structures = load_structures(workbook)
    if not structures:
        raise SystemExit(f"No structure rows found in {workbook}")

    to_screen = [s for s in structures if should_screen(s, pdf_folder)]
    skipped = [s for s in structures if s not in to_screen]
    for structure in skipped:
        print(f"Skipping {structure.number}: no new data")
    if not to_screen:
        print("Nothing new to screen.")
        return []

    outcomes: list[tuple[Structure, str, bool]] = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(viewport={"width": 1400, "height": 900})
        page = context.new_page()
        page.on("dialog", lambda dialog: dialog.accept())
        try:
            for structure in to_screen:
                print(f"Screening {structure.number} ...")
                result, _dest = screen_one(page, structure, pdf_folder)
                filing = requires_filing(result)
                write_result(workbook, structure, result, filing)
                outcomes.append((structure, result, filing))
                print(f"  {structure.number}: {result[:180]}")
                page.wait_for_timeout(2000)
        finally:
            context.close()
            browser.close()
    return outcomes


def main() -> None:
    workbook = Path(WORKBOOK_PATH)
    pdf_folder = Path(PDF_FOLDER)
    if not workbook.exists():
        raise SystemExit(f"Workbook not found: {workbook}")
    outcomes = run(workbook, pdf_folder)
    print("\nDone.")
    for structure, result, filing in outcomes:
        status = "NEED TO FILE" if filing else "FILE NOT REQUIRED"
        print(f"{structure.number}\t{status}\t{result[:120]}")


if __name__ == "__main__":
    main()
