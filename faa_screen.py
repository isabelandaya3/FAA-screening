"""Screen structures on the current FAA OE/AAA Pre-Screening Tool."""

from __future__ import annotations

import re
import shutil
import time
from pathlib import Path

from dataclasses import dataclass
from datetime import datetime
from hashlib import sha256
from math import ceil

import os
import subprocess

from openpyxl import load_workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeout
from playwright.sync_api import sync_playwright

from contextlib import contextmanager

@contextmanager
def _nullcontext():
    yield

# ---------------------------------------------------------------------------
# Edit these paths before you run the script.
# ---------------------------------------------------------------------------
WORKBOOK_PATH = r"C:\Users\And137460\OneDrive - Black & Veatch\PG&E Sacramento & LA OHTL - PGE Projects and Files\Projects\Sobrante\Working\30% Design\74066820 - Sobrante-Grizzly-Claremont #1 FAA Screening.xlsm"
PDF_FOLDER = r"C:\Users\And137460\OneDrive - Black & Veatch\PG&E Sacramento & LA OHTL - PGE Projects and Files\Projects\Sobrante\Working\30% Design\#1 FAA"
# ---------------------------------------------------------------------------

DOWNLOADS_DIR = ""
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
    used_elevation: int | None = None

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


def load_structures(workbook_path: Path): 
    wb = load_workbook(workbook_path, data_only=True)
    ws = wb[SHEET]
    rows: list[Structure] = []
    for r in range(2, ws.max_row + 1):
        number = ws[f"A{r}"].value
        if number is None or str(number).strip() == "":
            continue
        rows.append(
            Structure(
                row=r,
                number=str(number).strip(),
                pls_file=str(ws[f"B{r}"].value or "").strip(),
                lon_deg=_as_float(ws[f"C{r}"].value),
                lat_deg=_as_float(ws[f"D{r}"].value),
                elevation_ft=float(ws[f"E{r}"].value),
                height_ft=float(ws[f"F{r}"].value),
                lon_dms=str(ws[f"G{r}"].value).strip(),
                lat_dms=str(ws[f"H{r}"].value).strip(),
                traverseway=str(ws[f"I{r}"].value or "No Traverseway").strip(),
                on_airport=str(ws[f"J{r}"].value or "No").strip(),
                last_result=str(ws[f"K{r}"].value or "").strip(),
                last_hash=str(ws[f"M{r}"].value or "").strip(),
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
            str(structure.height_ft),
            str(structure.on_airport),
            str(structure.structure_type),
            datum,
            "USGS",
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
    ws[f"K{structure.row}"] = result_text
    ws[f"L{structure.row}"] = datetime.now()
    ws[f"M{structure.row}"] = input_fingerprint(structure, DATUM)

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
        data[f"D{next_row}"] = (
            structure.used_elevation
            if structure.used_elevation is not None
            else structure.elevation_whole
        )
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


def apply_nad83_datum(page) -> None:
    page.evaluate(
        """(datum) => {
            const el = document.querySelector('notice-criteria');
            const scope = angular.element(el).isolateScope() || angular.element(el).scope();
            const point = scope.data.points.find(p => p.include) || scope.data.points[0];
            point.datum = datum;
            if (scope.formData) scope.formData.datum = datum;
            if (scope.noticeData && scope.noticeData.verifyPointModalPoint) {
                scope.noticeData.verifyPointModalPoint.datum = datum;
            }
            scope.$apply();
        }""",
        DATUM,
    )


def read_accepted_elevation(page) -> tuple[int | None, str]:
    payload = page.evaluate(
        """() => {
            const el = document.querySelector('notice-criteria');
            if (!el || typeof angular === 'undefined') return null;
            const scope = angular.element(el).isolateScope() || angular.element(el).scope();
            if (!scope) return null;
            const point = scope.data.points.find(p => p.include) || scope.data.points[0];
            if (!point) return null;
            return {
                elev: point.siteElevation,
                source: point.siteElevationSource || ''
            };
        }"""
    )
    if not payload:
        return None, ""
    elev = payload.get("elev")
    try:
        elev_int = int(round(float(elev))) if elev is not None and elev != "" else None
    except (TypeError, ValueError):
        elev_int = None
    return elev_int, str(payload.get("source") or "")


def fill_point(page, structure: Structure) -> None:
    lat_body, lat_dir = parse_dms(structure.lat_dms)
    lon_body, lon_dir = parse_dms(structure.lon_dms)
    if not lat_dir:
        lat_dir = "N"
    if not lon_dir:
        lon_dir = "W"

    page.evaluate(
        """([lat, latDir, lon, lonDir, height, onAirport, datum]) => {
            if (!document.getElementById('terrain-working-indicator-notice')) {
                const marker = document.createElement('div');
                marker.id = 'terrain-working-indicator-notice';
                marker.style.display = 'none';
                document.body.appendChild(marker);
            }
            const el = document.querySelector('notice-criteria');
            const scope = angular.element(el).isolateScope() || angular.element(el).scope();
            const point = scope.data.points.find(p => p.include) || scope.data.points[0];
            point.formLat2 = lat;
            point.latDir = latDir;
            point.formLon2 = lon;
            point.lonDir = lonDir;
            point.datum = datum;
            point.structureHeight = Number(height);
            point.siteElevation = undefined;
            point.siteElevationSource = undefined;
            if (point.elevationDetails) {
                point.elevationDetails.comments = '';
            }
            point.state = 'COORD_DIRTY';
            scope.data.onAirport = !!onAirport;
            if (scope.formData) scope.formData.datum = datum;
            scope.$apply();
            return true;
        }""",
        [
            lat_body,
            lat_dir,
            lon_body,
            lon_dir,
            structure.height_whole,
            structure.on_airport.lower() == "yes",
            DATUM,
        ],
    )
    page.wait_for_timeout(800)
    validate = page.get_by_role("button", name="VALIDATE")
    validate.first.wait_for(state="visible", timeout=10000)
    validate.first.click()
    handle_point_modal(page, structure)
    apply_nad83_datum(page)
    if structure.lat_deg is not None and structure.lon_deg is not None:
        page.evaluate(
            """([lat, lon, datum]) => {
                const el = document.querySelector('notice-criteria');
                const scope = angular.element(el).isolateScope() || angular.element(el).scope();
                const point = scope.data.points.find(p => p.include) || scope.data.points[0];
                if (!point.lat) point.lat = lat;
                if (!point.lon) point.lon = lon;
                point.datum = datum;
                if (scope.formData) scope.formData.datum = datum;
                if (scope.isPointValid) scope.isPointValid(point);
                scope.$apply();
            }""",
            [structure.lat_deg, structure.lon_deg, DATUM],
        )
    page.wait_for_function(
        """() => {
            const el = document.querySelector('notice-criteria');
            const scope = angular.element(el).isolateScope() || angular.element(el).scope();
            const point = scope.data.points.find(p => p.include) || scope.data.points[0];
            return point && point.siteElevation !== undefined && point.siteElevation !== null
                && point.siteElevation !== '' && String(point.siteElevationSource || '').toUpperCase() !== 'USER';
        }""",
        timeout=10000,
    )
    elev, source = read_accepted_elevation(page)
    if elev is None or source.upper() == "USER":
        raise RuntimeError(
            f"FAA did not keep the NAD83/USGS site elevation (elev={elev!r}, source={source!r})"
        )
    structure.used_elevation = elev
    page.wait_for_timeout(400)


def handle_point_modal(page, structure: Structure) -> None:
    modal = page.locator("#verifyNoticePointModal")
    confirm = page.locator("#confirmPointModal")
    try:
        page.wait_for_function(
            """() => {
                const isOpen = (id) => {
                    const el = document.querySelector(id);
                    if (!el) return false;
                    const style = window.getComputedStyle(el);
                    return style.display !== 'none' && style.visibility !== 'hidden'
                        && (el.classList.contains('show') || el.classList.contains('in'));
                };
                return isOpen('#verifyNoticePointModal') || isOpen('#confirmPointModal');
            }""",
            timeout=25000,
        )
    except PlaywrightTimeout:
        raise RuntimeError(
            "FAA point validation dialog did not open after VALIDATE. "
            "The NAD83/USGS site elevation was not available."
        )
    if confirm.count() and confirm.first.is_visible() and not (
        modal.count() and modal.first.is_visible()
    ):
        page.get_by_role("button", name="Ok").click()
        page.wait_for_timeout(400)
        return
    page.wait_for_function(
        """() => {
            const el = document.querySelector('notice-criteria');
            if (!el || typeof angular === 'undefined') return false;
            const scope = angular.element(el).isolateScope() || angular.element(el).scope();
            const elev = scope && scope.noticeData && scope.noticeData.verifyPointModalPoint
                && scope.noticeData.verifyPointModalPoint.usgsSiteElev;
            return elev !== undefined && elev !== null && elev !== '';
        }""",
        timeout=20000,
    )
    keep = page.locator("#customRadioInline1")
    if keep.count():
        keep.check()
    page.evaluate(
        """(datum) => {
            const el = document.querySelector('notice-criteria');
            const scope = angular.element(el).isolateScope() || angular.element(el).scope();
            if (scope.noticeData && scope.noticeData.verifyPointModalPoint) {
                scope.noticeData.verifyPointModalPoint.siteElevationSource = 'USGS';
                scope.noticeData.verifyPointModalPoint.datum = datum;
            }
            if (scope.formData) scope.formData.datum = datum;
            scope.$apply();
        }""",
        DATUM,
    )
    page.get_by_role("button", name="Accept Point").click()
    page.wait_for_timeout(800)
    apply_nad83_datum(page)


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
    dest = dest.resolve()
    dest.parent.mkdir(parents=True, exist_ok=True)
    hide_modals(page)

    page.evaluate(
        """() => {
            window.close = function () {};
            window.print = function () {};
        }"""
    )

    print_btn = page.get_by_role("button", name="Print")
    print_btn.first.wait_for(state="visible", timeout=15000)

    # Capture the download event (we mainly want its URL)
    pdf_url = None
    download = None
    try:
        with page.expect_download(timeout=30000) as dl:
            print_btn.first.click()
        download = dl.value
        pdf_url = download.url
    except Exception as exc:
        print(f"  no download event ({exc}); will try response capture ...")

    saved = False

    # --- Attempt 1: normal save_as (works if browser survived) ---
    if download is not None:
        try:
            if dest.exists():
                dest.unlink()
            download.save_as(str(dest))
            saved = dest.exists() and dest.stat().st_size >= 1000
        except Exception as exc:
            print(f"  save_as failed ({exc}); fetching PDF by URL ...")

    # --- Attempt 2: re-fetch the PDF URL via the context API (survives page close) ---
    if not saved and pdf_url:
        try:
            resp = page.context.request.get(pdf_url, timeout=30000)
            if resp.ok:
                body = resp.body()
                if body and len(body) >= 1000:
                    if dest.exists():
                        dest.unlink()
                    with open(dest, "wb") as fh:
                        fh.write(body)
                    saved = True
        except Exception as exc:
            print(f"  URL fetch failed ({exc}); scanning downloads folder ...")

    # --- Attempt 3: last resort, newest file in the temp downloads dir ---
    if not saved:
        time.sleep(2)
        import glob
        candidates = sorted(
            glob.glob(os.path.join(DOWNLOADS_DIR, "**", "*"), recursive=True),
            key=lambda p: os.path.getmtime(p) if os.path.exists(p) else 0,
            reverse=True,
        )
        for cand in candidates:
            if os.path.isfile(cand) and os.path.getsize(cand) >= 1000:
                if dest.exists():
                    dest.unlink()
                shutil.copy2(cand, dest)
                saved = True
                break

    if not saved or not dest.exists() or dest.stat().st_size < 1000:
        raise RuntimeError(f"FAA Print PDF was not written: {dest}")
    print(f"  filed {dest} ({dest.stat().st_size} bytes)")

def screen_one(page, structure: Structure, pdf_folder: Path, booted: bool) -> tuple[str, Path]:
    try:
        def cold_boot():
            print("  cold-booting FAA app ...")
            page.goto("https://oeaaa.faa.gov/oeaaa/oe3a/main/", wait_until="domcontentloaded")
            page.wait_for_timeout(2000)
            dismiss_overlays(page)
            link = page.get_by_role("link", name="Pre-Screening Tool")
            if link.count():
                link.first.click()
            else:
                page.goto(PRESCREEN_URL, wait_until="domcontentloaded")

        def warm_nav():
            print("  reusing warm session ...")
            link = page.get_by_role("link", name="Pre-Screening Tool")
            if link.count():
                link.first.click()
            else:
                page.goto(PRESCREEN_URL, wait_until="domcontentloaded")

        # Try warm nav first (if booted); fall back to a full cold boot on failure
        rendered = False
        if booted:
            warm_nav()
            try:
                page.wait_for_selector("structure-type", timeout=8000)
                rendered = True
            except PlaywrightTimeout:
                print("  warm session stalled -> cold boot ...")

        if not rendered:
            for attempt in range(3):
                cold_boot()
                try:
                    page.wait_for_selector("structure-type", timeout=15000)
                    rendered = True
                    break
                except PlaywrightTimeout:
                    print(f"  structure-type not ready (attempt {attempt + 1}/3), retrying ...")
                    page.wait_for_timeout(2000)

        if not rendered:
            raise RuntimeError("Pre-Screening Tool did not load after fallback attempts")

        print("  selecting structure type ...")
        select_structure_type(page)
        print("  filling NAD83 point and keeping FAA/USGS site elevation ...")
        fill_point(page, structure)
        print(f"  using site elevation {structure.used_elevation} ft ...")
        print("  submitting ...")
        result = submit_and_read(page)
        if not RESULT_RE.search(result):
            raise RuntimeError(f"No FAA result text for {structure.number}: {result[:300]!r}")
        apply_nad83_datum(page)
        hide_modals(page)
        filing = requires_filing(result)
        dest = pdf_path(pdf_folder, structure.number, filing)
        remove_stale_pdfs(pdf_folder, structure.number, keep=None)
        print("  printing official FAA PDF ...")
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
    import tempfile
    global DOWNLOADS_DIR
    DOWNLOADS_DIR = tempfile.mkdtemp(prefix="faa_dl_")


    user_data_dir = r"C:\Users\And137460\FAA_edge_profile"

    # Kill any leftover automation Edge holding the profile lock
    lock = os.path.join(user_data_dir, "SingletonLock")
    if os.path.exists(lock):
        subprocess.run(
            ["taskkill", "/F", "/IM", "msedge.exe", "/T"],
            capture_output=True,
        )
        time.sleep(1.5)
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir,
            channel="msedge",
            headless=False,
            viewport={"width": 1400, "height": 900},
            chromium_sandbox=True,
            accept_downloads=True,
            downloads_path=DOWNLOADS_DIR,
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.on("dialog", lambda dialog: dialog.accept())
        try:
            booted = False
            for structure in to_screen:
                print(f"Screening {structure.number} ...")
                result, _dest = screen_one(page, structure, pdf_folder, booted)
                booted = True
                filing = requires_filing(result)
                write_result(workbook, structure, result, filing)
                outcomes.append((structure, result, filing))
                print(f"  {structure.number}: {result[:180]}")
                page.wait_for_timeout(2000)
        finally:
            context.close()
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
