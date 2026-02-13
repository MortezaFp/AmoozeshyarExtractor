import importlib.util
import re
import subprocess
import sys
import time
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_NAME = "Amoozeshyar Offered Courses Extractor"

RAW_EXCEL_NAME = "لیست دروس ارائه شده آموزشیار.xlsx"
SPECIALIZED_EXCEL_NAME = "لیست دروس تخصصی.xlsx"
GENERAL_EXCEL_NAME = "لیست دروس عمومی.xlsx"


REQUIRED_PACKAGES = [
    "playwright",
    "pandas",
    "openpyxl",
    "reportlab",
    "arabic-reshaper",
    "python-bidi",
]


def ensure_package(package_name: str) -> None:
    if importlib.util.find_spec(package_name) is None:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])


for pkg in REQUIRED_PACKAGES:
    ensure_package(pkg)


try:
    playwright_sync_api = importlib.import_module("playwright.sync_api")
    PlaywrightError = playwright_sync_api.Error
    PlaywrightTimeoutError = playwright_sync_api.TimeoutError
    sync_playwright = playwright_sync_api.sync_playwright
except Exception as exc:
    raise RuntimeError(
        "Could not import Playwright. Try: pip install playwright ; python -m playwright install chrome"
    ) from exc


pd = importlib.import_module("pandas")
bidi_algorithm = importlib.import_module("bidi.algorithm")
get_display = bidi_algorithm.get_display
arabic_reshaper = importlib.import_module("arabic_reshaper")

colors = importlib.import_module("reportlab.lib.colors")
pagesizes = importlib.import_module("reportlab.lib.pagesizes")
A4 = pagesizes.A4
landscape = pagesizes.landscape
getSampleStyleSheet = importlib.import_module(
    "reportlab.lib.styles"
).getSampleStyleSheet
ParagraphStyle = importlib.import_module("reportlab.lib.styles").ParagraphStyle
TA_CENTER = importlib.import_module("reportlab.lib.enums").TA_CENTER
mm = importlib.import_module("reportlab.lib.units").mm
pdfmetrics = importlib.import_module("reportlab.pdfbase.pdfmetrics")
TTFont = importlib.import_module("reportlab.pdfbase.ttfonts").TTFont
platypus = importlib.import_module("reportlab.platypus")
LongTable = platypus.LongTable
Paragraph = platypus.Paragraph
SimpleDocTemplate = platypus.SimpleDocTemplate
Spacer = platypus.Spacer
TableStyle = platypus.TableStyle


TARGET_URL = (
    "https://eserv.iau.ir/EServices/handleCourseClassSearchAction.do"
    "?parameter%28menuItem%29=0_0"
    "&dispatch=selectStudentParameter"
    "&subject=CourseClass"
    "&editable=false"
    "&previewable=false"
    "&parameter%28f%5EtermRef%29=%24%7BuserProperty%28operationalTerm.id%29%7D"
    "&addable=false"
    "&refParameter%28selectedText%29=parameter%28courseClassText%29"
    "&form=CourseClassList2student"
    "&selection=0"
    "&parameter%28groupIndex%29=0"
    "&deleteable=false"
    "&parameter%28f%5EstudentRef%29=%24%7BuserProperty%28studentDto.id%29%7D"
    "&parameter%28finder%29=findCourseClass4Student"
    "&refParameter%28selectedId%29=parameter%28courseClassRef%29"
    "&parameter(menuItem)=0_0"
    "&parameter(groupIndex)=0"
    "&__rp=898686051"
    "&menuGroup=Planning"
    "&menuItemName=StudentCourseClassAllSearch"
    "&_H0__=-153"
    "&_H2__=18"
    "&_H1__=1386"
)
START_URL = "https://eserv.iau.ir/EServices/startAction.do"

MEANINGFUL_COLUMNS = [
    "كد درس",
    "نام درس",
    "نوع درس",
    "تعداد واحد نظري",
    "تعداد واحد عملي",
    "كد ارائه کلاس درس",
    "نام كلاس درس",
    "زمانبندي تشکيل کلاس",
    "استاد",
    "ساير اساتيد",
    "حداكثر ظرفيت",
    "تعداد ثبت نامي تاکنون",
    "زمان امتحان",
    "مكان برگزاري",
    "مقطع ارائه درس",
    "نوع ارائه",
    "سطح ارائه",
    "دانشجويان مجاز به اخذ کلاس",
    "گروه آموزشی",
    "دانشکده",
    "واحد",
    "استان",
]

GROUP_LEVEL_TEXT = "ارائه در سطح گروه آموزشی"
FACULTY_LEVEL_TEXT = "ارائه در سطح دانشکده"

FONT_CANDIDATES = [
    SCRIPT_DIR / "B_Nazanin_Bold.ttf",
    Path(r"C:\Windows\Fonts\arial.ttf"),
    Path(r"C:\Windows\Fonts\tahoma.ttf"),
]


def normalize_text(value) -> str:
    if value is None:
        return ""
    if pd.isna(value):
        return ""
    return str(value).strip().replace("_", "-")


def shape_persian(value) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    reshaped = arabic_reshaper.reshape(text)
    return get_display(reshaped)


PERSIAN_ALPHA_ORDER = {
    char: index
    for index, char in enumerate(
        "اآبپتثجچحخدذرزژسشصضطظعغفقکگلمنوهی",
        start=1,
    )
}


def normalize_persian_for_sort(value) -> str:
    text = normalize_text(value)
    if not text:
        return ""

    replacements = {
        "ي": "ی",
        "ى": "ی",
        "ئ": "ی",
        "ك": "ک",
        "ة": "ه",
        "ۀ": "ه",
        "ؤ": "و",
        "أ": "ا",
        "إ": "ا",
        "ٱ": "ا",
        "آ": "ا",
        "۰": "0",
        "۱": "1",
        "۲": "2",
        "۳": "3",
        "۴": "4",
        "۵": "5",
        "۶": "6",
        "۷": "7",
        "۸": "8",
        "۹": "9",
        "٠": "0",
        "١": "1",
        "٢": "2",
        "٣": "3",
        "٤": "4",
        "٥": "5",
        "٦": "6",
        "٧": "7",
        "٨": "8",
        "٩": "9",
        "\u200c": " ",
        "\u200f": "",
        "\u200e": "",
    }
    for src, dst in replacements.items():
        text = text.replace(src, dst)

    text = re.sub(r"[\u064B-\u065F\u0670\u06D6-\u06ED]", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def persian_sort_key(value) -> str:
    text = normalize_persian_for_sort(value)
    if not text:
        return ""

    chunks = []
    for char in text:
        if char.isdigit():
            chunks.append(f"0{int(char):03d}")
            continue
        rank = PERSIAN_ALPHA_ORDER.get(char)
        if rank is not None:
            chunks.append(f"1{rank:03d}")
        else:
            chunks.append(f"9{ord(char):04d}")
    return "".join(chunks)


def normalize_header_key(value) -> str:
    return normalize_persian_for_sort(value).replace(" ", "")


def reverse_dataframe_columns(df):
    columns = list(df.columns)
    if len(columns) <= 1:
        return df
    return df[columns[::-1]].copy()


def rtl_paragraph(
    value,
    style,
    wrap_chars: int | None = None,
    reverse_visual_lines: bool = False,
) -> Paragraph:
    text = normalize_text(value)
    if not text:
        return Paragraph("", style)

    split_lines = [line for line in text.splitlines() if line.strip()]
    if not split_lines:
        split_lines = [text]

    lines = []
    for line in split_lines:
        if wrap_chars and len(line) > wrap_chars:
            parts = [
                p
                for p in re.findall(rf".{{1,{wrap_chars}}}(?:\s+|$)", line + " ")
                if p.strip()
            ]
            if parts:
                lines.extend([p.strip() for p in parts])
            else:
                lines.append(line)
        else:
            lines.append(line)

    shaped_lines = [shape_persian(line).strip() for line in lines if line]
    if not shaped_lines:
        return Paragraph("", style)

    if reverse_visual_lines and len(shaped_lines) > 1:
        shaped_lines = list(reversed(shaped_lines))

    paragraph_text = "<br/>".join(
        line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        for line in shaped_lines
    )
    return Paragraph(paragraph_text, style)


def register_font() -> str:
    candidates = list(FONT_CANDIDATES)
    user_fonts = Path.home() / "AppData" / "Local" / "Microsoft" / "Windows" / "Fonts"
    if user_fonts.exists():
        candidates.extend(sorted(user_fonts.glob("*Nazanin*.ttf")))
        candidates.extend(sorted(user_fonts.glob("*nazanin*.ttf")))

    for font_path in candidates:
        if font_path.exists():
            font_name = f"CustomFont_{font_path.stem}"
            pdfmetrics.registerFont(TTFont(font_name, str(font_path)))
            return font_name

    raise RuntimeError("No suitable Persian-supporting font found.")


def dataframe_to_pdf(
    df,
    pdf_path: Path,
    title: str,
    font_name: str,
    header_bg_color,
    stripe_bg_color,
    grid_color,
) -> None:
    df = df.fillna("")

    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=landscape(A4),
        leftMargin=5 * mm,
        rightMargin=5 * mm,
        topMargin=6 * mm,
        bottomMargin=6 * mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "CustomTitle",
        parent=styles["Title"],
        fontName=font_name,
        fontSize=16,
        alignment=TA_CENTER,
    )

    body_style = ParagraphStyle(
        "CustomBody",
        parent=styles["BodyText"],
        fontName=font_name,
        fontSize=8,
        leading=10,
        alignment=TA_CENTER,
        leftIndent=0,
        rightIndent=0,
        firstLineIndent=0,
        spaceBefore=0,
        spaceAfter=0,
        wordWrap="RTL",
    )

    header_style = ParagraphStyle(
        "CustomHeader",
        parent=body_style,
        textColor=colors.white,
        wordWrap="LTR",
    )

    problematic_headers = {
        normalize_header_key("كد ارائه کلاس درس"),
        normalize_header_key("کد ارائه کلاس درس"),
        normalize_header_key("زمانبندي تشکيل کلاس"),
        normalize_header_key("زمانبندی تشکیل کلاس"),
    }

    table_data = []
    header_row = [
        rtl_paragraph(
            col,
            header_style,
            wrap_chars=10 if normalize_header_key(col) in problematic_headers else None,
            reverse_visual_lines=False,
        )
        for col in df.columns
    ]
    table_data.append(header_row)

    wrap_chars = 14 if len(df.columns) >= 10 else 24
    for _, row in df.iterrows():
        row_vals = [
            rtl_paragraph(row[col], body_style, wrap_chars) for col in df.columns
        ]
        table_data.append(row_vals)

    usable_width = landscape(A4)[0] - 10 * mm
    table_width = usable_width * 0.995
    col_width = table_width / max(1, len(df.columns))
    col_widths = [col_width] * len(df.columns)

    table = LongTable(table_data, colWidths=col_widths, repeatRows=1)
    table.hAlign = "CENTER"
    table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), font_name),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BACKGROUND", (0, 0), (-1, 0), header_bg_color),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.25, grid_color),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, stripe_bg_color]),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )

    elements = [Paragraph(shape_persian(title), title_style), Spacer(1, 3 * mm), table]
    doc.build(elements)


def postprocess_excel_to_pdfs(
    source_excel: Path,
) -> tuple[Path, Path, Path, Path, int, int]:
    df = pd.read_excel(source_excel)

    existing = [col for col in MEANINGFUL_COLUMNS if col in df.columns]
    if len(existing) < 16:
        raise RuntimeError(
            "Input Excel does not look like expected course export columns."
        )
    df = df[existing].copy()

    if "دانشکده" not in df.columns or "سطح ارائه" not in df.columns:
        raise RuntimeError("Required columns دانشکده or سطح ارائه are missing.")

    filtered = df[df["دانشکده"].astype(str).str.contains("143", na=False)].copy()
    group_df = filtered[
        filtered["سطح ارائه"].astype(str).str.strip() == GROUP_LEVEL_TEXT
    ].copy()
    faculty_df = filtered[
        filtered["سطح ارائه"].astype(str).str.strip() == FACULTY_LEVEL_TEXT
    ].copy()

    base_cols = [
        "كد درس",
        "نام درس",
        "نوع درس",
        "تعداد واحد نظري",
        "تعداد واحد عملي",
        "كد ارائه کلاس درس",
        "نام كلاس درس",
        "زمانبندي تشکيل کلاس",
        "استاد",
        "ساير اساتيد",
        "حداكثر ظرفيت",
        "تعداد ثبت نامي تاکنون",
        "زمان امتحان",
        "مكان برگزاري",
        "مقطع ارائه درس",
        "نوع ارائه",
    ]
    keep_after_drop = [
        col
        for col in base_cols
        if col not in {"ساير اساتيد", "تعداد ثبت نامي تاکنون", "نوع ارائه"}
    ]

    group_df = group_df[[c for c in keep_after_drop if c in group_df.columns]].copy()
    faculty_df = faculty_df[
        [c for c in keep_after_drop if c in faculty_df.columns]
    ].copy()

    if "نام درس" in group_df.columns:
        group_df = group_df.sort_values(
            by="نام درس", kind="stable", key=lambda s: s.map(persian_sort_key)
        )
    if "نام درس" in faculty_df.columns:
        faculty_df = faculty_df.sort_values(
            by="نام درس", kind="stable", key=lambda s: s.map(persian_sort_key)
        )

    group_df = group_df.fillna("").reset_index(drop=True)
    faculty_df = faculty_df.fillna("").reset_index(drop=True)

    group_df = reverse_dataframe_columns(group_df)
    faculty_df = reverse_dataframe_columns(faculty_df)

    font_name = register_font()
    out_dir = SCRIPT_DIR
    group_excel = out_dir / SPECIALIZED_EXCEL_NAME
    faculty_excel = out_dir / GENERAL_EXCEL_NAME
    group_pdf = out_dir / "لیست دروس تخصصی.pdf"
    faculty_pdf = out_dir / "لیست دروس عمومی.pdf"

    group_df.to_excel(group_excel, index=False)
    faculty_df.to_excel(faculty_excel, index=False)

    green_header = colors.HexColor("#14532d")
    green_stripe = colors.HexColor("#dcfce7")
    blue_header = colors.HexColor("#1e3a8a")
    blue_stripe = colors.HexColor("#dbeafe")
    dark_gray_grid = colors.HexColor("#374151")

    dataframe_to_pdf(
        group_df,
        group_pdf,
        "لیست دروس تخصصی",
        font_name,
        green_header,
        green_stripe,
        dark_gray_grid,
    )
    dataframe_to_pdf(
        faculty_df,
        faculty_pdf,
        "لیست دروس عمومی",
        font_name,
        blue_header,
        blue_stripe,
        dark_gray_grid,
    )

    return (
        group_excel,
        faculty_excel,
        group_pdf,
        faculty_pdf,
        len(group_df),
        len(faculty_df),
    )


EXTRACT_TABLE_JS = r"""
() => {
    const clean = (txt) => (txt || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
    const norm = (s) => clean(s).replace(/ي/g, 'ی').replace(/ك/g, 'ک');

    const meaningful = [
        'كد درس','نام درس','نوع درس','تعداد واحد نظري','تعداد واحد عملي','كد ارائه کلاس درس',
        'نام كلاس درس','زمانبندي تشکيل کلاس','استاد','ساير اساتيد','حداكثر ظرفيت','تعداد ثبت نامي تاکنون',
        'زمان امتحان','مكان برگزاري','مقطع ارائه درس','نوع ارائه','سطح ارائه','دانشجويان مجاز به اخذ کلاس',
        'گروه آموزشی','دانشکده','واحد','استان'
    ];
    const meaningfulNorm = meaningful.map(norm);

    let best = null;
    const tables = [...document.querySelectorAll('table')];
    for (const table of tables) {
        const rows = [...table.querySelectorAll('tr')];
        for (const row of rows) {
            const cells = [...row.querySelectorAll('th,td')].map((c) => norm(c.innerText || c.textContent || ''));
            if (cells.length < 12) continue;
            const matchCount = meaningfulNorm.filter((h) => cells.includes(h)).length;
            if (matchCount < 10) continue;
            const score = matchCount * 100 - Math.abs(cells.length - 23);
            if (!best || score > best.score) {
                best = { table, row, cells, score };
            }
        }
    }

    if (!best) return { headers: [], rows: [], pageInfo: '' };

    const rawHeadersNorm = best.cells;
    const rawHeadersText = [...best.row.querySelectorAll('th,td')].map((c) => clean(c.innerText || c.textContent || ''));
    const indexMap = new Map();
    for (let i = 0; i < rawHeadersNorm.length; i++) {
        if (!indexMap.has(rawHeadersNorm[i])) indexMap.set(rawHeadersNorm[i], i);
    }

    const headers = [];
    const headerIndexes = [];
    for (let i = 0; i < meaningfulNorm.length; i++) {
        const n = meaningfulNorm[i];
        if (indexMap.has(n)) {
            headerIndexes.push(indexMap.get(n));
            headers.push(rawHeadersText[indexMap.get(n)] || meaningful[i]);
        }
    }

    if (headers.length < 10) return { headers: [], rows: [], pageInfo: '' };

    const dataRows = [];
    const allRows = [...best.table.querySelectorAll('tr')];
    const headerPos = allRows.indexOf(best.row);

    for (let i = headerPos + 1; i < allRows.length; i++) {
        const tr = allRows[i];
        const cellsRaw = [...tr.querySelectorAll('td')].map((c) => clean(c.innerText || c.textContent || ''));
        if (cellsRaw.length < headerIndexes.length) continue;
        if (cellsRaw.some((v) => /نتايج\s*جستجو|کلیه\s*حقوق/i.test(v))) continue;

        const hasSerial = cellsRaw.some((v) => /^\d+$/.test(v || ''));
        const hasCourseCode = cellsRaw.some((v) => /^\d{4,}$/.test(v || ''));
        if (!hasSerial || !hasCourseCode) continue;

        const obj = {};
        let nonEmpty = 0;
        for (let c = 0; c < headerIndexes.length; c++) {
            const idx = headerIndexes[c];
            const key = headers[c];
            const value = cellsRaw[idx] || '';
            obj[key] = value;
            if (value) nonEmpty += 1;
        }
        if (nonEmpty >= 4) dataRows.push(obj);
    }

	const bodyText = clean(document.body ? document.body.innerText : '');
	const pageInfoMatch = bodyText.match(/نتايج\s*جستجو\s*\(\s*ركورد\s*\d+\s*تا\s*\d+\s*از\s*\d+\s*ركورد\s*\)/);
	const pageInfo = pageInfoMatch ? pageInfoMatch[0] : '';

	return { headers, rows: dataRows, pageInfo };
}
"""


NEXT_PAGE_DISABLED_JS = r"""
() => {
	const candidates = [...document.querySelectorAll('input,button,a')];
	for (const el of candidates) {
		const t = (el.innerText || el.textContent || el.value || '').trim();
		if (t.includes('صفحه بعد')) {
			const disabled = el.disabled || el.getAttribute('aria-disabled') === 'true' || el.classList.contains('disabled');
			return !!disabled;
		}
	}
	return true;
}
"""


CLICK_NEXT_PAGE_JS = r"""
() => {
	const candidates = [...document.querySelectorAll('input,button,a')];
	for (const el of candidates) {
		const t = (el.innerText || el.textContent || el.value || '').trim();
		if (t.includes('صفحه بعد')) {
			if (el.disabled) return false;
			el.click();
			return true;
		}
	}
	return false;
}
"""


def wait_for_login(page) -> None:
    safe_goto(page, "https://eserv.iau.ir")
    print("\nLogin in the opened Chrome window, then press Enter here...")
    input()


def safe_goto(page, url: str, timeout_ms: int = 45000) -> None:
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
        return
    except PlaywrightError as exc:
        message = str(exc)
        interrupted = "interrupted by another navigation" in message
        same_target = url in message or url.rstrip("/") in (page.url or "").rstrip("/")
        if interrupted and same_target:
            try:
                page.wait_for_load_state("domcontentloaded", timeout=timeout_ms)
            except PlaywrightTimeoutError:
                pass
            return
        raise


def is_session_expired(page) -> bool:
    url = (page.url or "").lower()
    if "loginpage.jsp" in url or "/login" in url:
        return True
    try:
        login_inputs = page.locator("input[type='password'], input[name*='password']")
        if login_inputs.count() > 0:
            return True
    except Exception:
        pass
    return False


def get_menu_frame(page, timeout_ms: int = 15000):
    deadline = time.time() + (timeout_ms / 1000)
    while time.time() < deadline:
        for frame in page.frames:
            frame_url = (frame.url or "").lower()
            if "cache?a=menu" in frame_url:
                return frame
        page.wait_for_timeout(250)
    return None


def open_course_search_from_menu(page) -> bool:
    # Retry because menu iframe and items can load asynchronously.
    for _ in range(6):
        menu_frame = get_menu_frame(page, timeout_ms=4000)
        if menu_frame is None:
            continue

        try:
            planning_section = menu_frame.get_by_text(
                "برنامه ريزي آموزشي نيمسال تحصيلي", exact=False
            )
            if planning_section.count() > 0:
                planning_section.first.click(timeout=3000)
                page.wait_for_timeout(500)
        except Exception:
            pass

        try:
            course_search = menu_frame.get_by_text(
                "جستجوي كلاس درسهای ارائه شده", exact=False
            )
            if course_search.count() > 0:
                course_search.first.click(timeout=4000)
                page.wait_for_timeout(1800)
                if is_on_course_search_page(page) or wait_for_search_controls(
                    page, 6000
                ):
                    return True
        except Exception:
            pass

        page.wait_for_timeout(600)

    return False


def is_on_course_search_page(page) -> bool:
    try:
        url = (page.url or "").lower()
        if "courseclass" in url or "psearchaction.do" in url:
            return True

        has_controls = page.evaluate(
            r"""
            () => {
                const submit = document.querySelector('#submitBtn');
                const rowCount = document.querySelector("select[name='parameter(rowCount)']");
                const hasCourseSearchTitle = (document.body?.innerText || '').includes('جستجوي كلاس درس');
                return !!submit || !!rowCount || hasCourseSearchTitle;
            }
            """
        )
        return bool(has_controls)
    except Exception:
        return False


def wait_for_search_controls(page, timeout_ms: int) -> bool:
    deadline = time.time() + (timeout_ms / 1000)
    while time.time() < deadline:
        try:
            if is_on_course_search_page(page):
                return True
            if page.locator("#submitBtn").count() > 0:
                return True
            if page.locator("button:has-text('جستجو')").count() > 0:
                return True
            if page.locator("input[value='جستجو']").count() > 0:
                return True
        except Exception:
            pass
        page.wait_for_timeout(250)
    return False


def force_open_course_search(page) -> None:
    # Prefer dashboard menu navigation (more stable than direct deep-link URL).
    safe_goto(page, START_URL)

    if open_course_search_from_menu(page):
        try:
            page.wait_for_timeout(2000)
            if is_on_course_search_page(page):
                return
            if wait_for_search_controls(page, 12000):
                return
            return
        except PlaywrightTimeoutError:
            pass

    try:
        if is_on_course_search_page(page):
            return
        if wait_for_search_controls(page, 5000):
            return
        return
    except PlaywrightTimeoutError:
        pass

    # Fallback if menu route fails.
    safe_goto(page, TARGET_URL)

    if open_course_search_from_menu(page):
        page.wait_for_timeout(1800)
    if not is_on_course_search_page(page):
        raise RuntimeError(
            "Could not open course search page via menu or fallback URL."
        )


def get_result_summary(page) -> str:
    try:
        summary = page.evaluate(
            r"""
            () => {
                const t = (document.body?.innerText || '');
                const m = t.match(/نتايج\s*جستجو\s*\([^\)]*\)/);
                return m ? m[0] : '';
            }
            """
        )
        return (summary or "").strip()
    except Exception:
        return ""


def click_search_button(page) -> bool:
    for selector in ["#submitBtn", "button:has-text('جستجو')", "input[value='جستجو']"]:
        try:
            loc = page.locator(selector)
            if loc.count() > 0:
                loc.first.click()
                return True
        except Exception:
            pass

    try:
        return bool(
            page.evaluate(
                r"""
                () => {
                    const btn = document.querySelector('#submitBtn');
                    if (btn) { btn.click(); return true; }
                    const byValue = document.querySelector("input[value='جستجو']");
                    if (byValue) { byValue.click(); return true; }
                    return false;
                }
                """
            )
        )
    except Exception:
        return False


def ensure_row_count_100(page) -> str:
    last_summary = ""
    for _ in range(5):
        try:
            row_count_select = page.locator("select[name='parameter(rowCount)']")
            if row_count_select.count() > 0:
                row_count_select.first.select_option("100")
        except Exception:
            pass

        if not click_search_button(page):
            page.wait_for_timeout(800)
            continue

        page.wait_for_timeout(2200)
        last_summary = get_result_summary(page)
        if re.search(r"ركورد\s*\d+\s*تا\s*100\s*از", last_summary):
            return last_summary

        try:
            page.evaluate(
                r"""
                () => {
                    const select = document.querySelector("select[name='parameter(rowCount)']");
                    if (select) {
                        select.value = '100';
                        select.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                }
                """
            )
        except Exception:
            pass

    return last_summary


def set_row_count_100_and_search(page) -> None:
    ensure_row_count_100(page)


def wait_for_results(page) -> None:
    for _ in range(2):
        force_open_course_search(page)
        if is_on_course_search_page(page):
            break
        page.wait_for_timeout(1500)

    if is_session_expired(page):
        print("Session expired. Please login again, then press Enter...")
        input()
        force_open_course_search(page)

    if not is_on_course_search_page(page):
        try:
            if not wait_for_search_controls(page, 30000):
                raise PlaywrightTimeoutError("Timed out waiting for search controls")
        except PlaywrightTimeoutError:
            raise RuntimeError("Could not open course search page (جستجوي كلاس درس).")

    if not is_on_course_search_page(page):
        raise RuntimeError("Could not open course search page (جستجوي كلاس درس).")

    search_clicked = click_search_button(page)

    if not search_clicked:
        raise RuntimeError("Could not find/click search button (جستجو) on the page.")

    page.wait_for_timeout(2000)

    # Force 100 rows per page and verify.
    summary = ensure_row_count_100(page)

    print(f"Result summary after rowCount=100: {summary or 'N/A'}")
    if not re.search(r"ركورد\s*\d+\s*تا\s*100\s*از", summary):
        raise RuntimeError(
            f"Could not set row count to 100. Current summary: {summary or 'N/A'}"
        )


def scrape_all_pages(page) -> list[dict]:
    collected_rows: list[dict] = []
    seen_pages: set[str] = set()
    empty_pages = 0

    while True:
        if is_session_expired(page):
            print(
                "Session expired during scraping. Stopping and saving collected rows."
            )
            break

        extracted = page.evaluate(EXTRACT_TABLE_JS)
        page_info = (extracted.get("pageInfo") or "").strip()
        rows = extracted.get("rows") or []

        if page_info and page_info in seen_pages:
            break
        if page_info:
            seen_pages.add(page_info)

        collected_rows.extend(rows)
        print(f"Collected rows: {len(collected_rows)}")

        if rows:
            empty_pages = 0
        else:
            empty_pages += 1
            if empty_pages >= 3:
                raise RuntimeError(
                    "No rows extracted for 3 consecutive pages. Stopping to avoid bad export."
                )

        is_disabled = bool(page.evaluate(NEXT_PAGE_DISABLED_JS))
        if is_disabled:
            break

        clicked = bool(page.evaluate(CLICK_NEXT_PAGE_JS))
        if not clicked:
            break

        page.wait_for_timeout(1800)

    return collected_rows


def save_excel(rows: list[dict]) -> Path:
    output_path = SCRIPT_DIR / RAW_EXCEL_NAME
    if not rows:
        pd.DataFrame([{"message": "No rows found"}]).to_excel(output_path, index=False)
    else:
        df = pd.DataFrame(rows)

        def norm_col(name: str) -> str:
            return (
                re.sub(r"\s+", " ", str(name).strip())
                .replace("ي", "ی")
                .replace("ك", "ک")
            )

        norm_to_real = {norm_col(c): c for c in df.columns}
        ordered_cols = []
        for wanted in MEANINGFUL_COLUMNS:
            key = norm_col(wanted)
            if key in norm_to_real:
                ordered_cols.append(norm_to_real[key])

        if ordered_cols:
            df = df[ordered_cols]

        df = reverse_dataframe_columns(df)

        df.to_excel(output_path, index=False)
    return output_path


def main() -> None:
    print(f"Starting {PROJECT_NAME}...")
    print("If first run fails, execute once: python -m playwright install chrome")

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="chrome", headless=False)
        context = browser.new_context()
        page = context.new_page()

        wait_for_login(page)
        wait_for_results(page)
        rows = scrape_all_pages(page)
        excel_file = save_excel(rows)
        (
            group_excel,
            faculty_excel,
            group_pdf,
            faculty_pdf,
            group_count,
            faculty_count,
        ) = postprocess_excel_to_pdfs(excel_file)

        print(f"\nDone. Exported {len(rows)} rows to: {excel_file}")
        print(f"Specialized Excel: {group_excel}")
        print(f"General Excel: {faculty_excel}")
        print(f"Specialized rows: {group_count} -> {group_pdf}")
        print(f"General rows: {faculty_count} -> {faculty_pdf}")
        print("Browser stays open for review. Press Enter to close.")
        input()

        context.close()
        browser.close()


if __name__ == "__main__":
    main()
