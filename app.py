"""Куратор төлем есебі — сайт. Іске қосу: streamlit run app.py
Екі рөл: АДМИН (ADMIN_PASSWORD) — барлық функция, отчеттер, суммалар; ҚЫЗМЕТКЕР (STAFF_PASSWORD) — тек негізгі бет."""
import os, json, tempfile
from datetime import datetime, date, timedelta
from pathlib import Path
import pandas as pd
import streamlit as st
import altair as alt
from dotenv import load_dotenv
import core
import reports
import external as ext

load_dotenv()
ROSTER = Path(__file__).parent / "data" / "roster.xlsx"
ADMIN_PW = os.getenv("ADMIN_PASSWORD", "")
STAFF_PW = os.getenv("STAFF_PASSWORD", "") or os.getenv("APP_PASSWORD", "")
REPORTS_PW = os.getenv("REPORTS_PASSWORD", "")      # «Отчеттер» бөлімдеріне қосымша құпиясөз (бос болса — қосымша құпиясөз жоқ)
DEF_SHEETS = os.getenv("UP_SHEETS", "қазан, қыркүйек")
DEF_URL = os.getenv("UP_URL", "")
SETTINGS = Path(__file__).parent / "data" / "settings.json"        # жоспар сандары (СЫРТҚЫ отчет)
PAY_FILE = Path(__file__).parent / "data" / "payments.xlsx"     # платформадан төлемдер экспорты (СЫРТҚЫ)
ROSTER_URL = os.getenv("ROSTER_URL", "")          # куратор тізімі Google Sheets-те болса (сервердегі жариялауға ыңғайлы)
PAYMENTS_URL = os.getenv("PAYMENTS_URL", "")      # платформа төлемдері Google Sheets-те болса


@st.cache_data(ttl=600, show_spinner="Деректер жүктелуде…")
def sync_remote(roster_url: str, pay_url: str) -> list:
    """Сілтеме берілсе, куратор тізімі мен платформа төлемдерін Google Sheets-тен data/ папкасына жүктейді.
    Сондықтан оқушылардың деректері GitHub-қа/сайт файлдарына салынбайды."""
    errs = []
    for url, dest, name in ((roster_url, ROSTER, "куратор тізімі"), (pay_url, PAY_FILE, "платформа төлемдері")):
        if url:
            try:
                dest.parent.mkdir(exist_ok=True)
                core.fetch_google_sheet(url, dest)
            except core.UpFileError as e:
                errs.append(f"{name}: {e}")
    return errs

st.set_page_config(page_title="JUZ40 · Келесі айға төлем", page_icon="📊", layout="wide")
import base64

RED, RED_D, RED_L = "#b3262e", "#6b1218", "#e5484d"
SERIES = ["#e5484d", "#f2a0a5", "#9a9aa2", "#b08b8e", "#8f1d24", "#5d5d63"]     # сызықтарға: қызыл-ақ-сұр
LOGO = Path(__file__).parent / "assets" / "logo.png"


def logo_img(height: int = 44) -> str:
    try:
        b = base64.b64encode(LOGO.read_bytes()).decode()
        return f'<img src="data:image/png;base64,{b}" style="height:{height}px;border-radius:8px;display:block">'
    except Exception:
        return '<span style="font-weight:800;letter-spacing:2px">JUZ40</span>'


CSS = """
<style>
.stApp{background:radial-gradient(1100px 520px at 88% -8%,@g1@,transparent 60%),
       radial-gradient(800px 400px at -5% 105%,@g2@,transparent 60%),@bg@;}
[data-testid="stSidebar"]{background:@side@;border-right:1px solid @bd@}
[data-testid="stHeader"]{background:transparent}
footer,#MainMenu,[data-testid="stAppDeployButton"]{visibility:hidden}
h1,h2,h3{letter-spacing:.2px}
h2,h3{border-left:4px solid #b3262e;padding-left:.6rem}
[data-testid="stMetric"]{background:linear-gradient(150deg,@card1@,@card2@);border:1px solid @bd@;
       border-left:4px solid #b3262e;border-radius:14px;padding:14px 16px}
button[data-baseweb="tab"]{font-weight:600}
button[data-baseweb="tab"][aria-selected="true"]{color:#b3262e}
[data-baseweb="tab-highlight"]{background:#b3262e!important}
.stButton>button,.stDownloadButton>button,[data-testid="stFormSubmitButton"]>button{border:1px solid #b3262e;
       background:linear-gradient(180deg,#7d1a21,#5e0f15);color:#fff;border-radius:10px;font-weight:600}
.stButton>button:hover,.stDownloadButton>button:hover,[data-testid="stFormSubmitButton"]>button:hover{
       border-color:#ff6b73;background:linear-gradient(180deg,#9a222a,#6b1218);color:#fff}
[data-testid="stDataFrame"]{border:1px solid @bd@;border-radius:12px;overflow:hidden}
code,pre{border-radius:10px}
.hero{display:flex;align-items:center;gap:18px;padding:16px 22px;margin:-8px 0 18px;border-radius:16px;
      background:linear-gradient(100deg,#6b1218 0%,#3a0a0e 55%,#161011 100%);border:1px solid #4a1a1f}
.hero .badge{margin-left:auto;background:rgba(0,0,0,.35);border:1px solid #7d1a21;
      padding:6px 14px;border-radius:999px;font-weight:600;font-size:.9rem;white-space:nowrap}
.login-hero{text-align:center;margin:8vh 0 14px}
.login-hero .logo{display:inline-block;padding:0;border-radius:14px;box-shadow:0 10px 40px rgba(179,38,46,.35)}
.login-hero h1{margin:18px 0 4px;font-size:2.2rem;border:0!important;padding:0!important}
[data-testid="stForm"]{background:linear-gradient(160deg,@form1@,@form2@);border:1px solid @bd@;border-radius:18px;
      padding:22px 22px 14px;box-shadow:0 12px 40px @shadow@}
/* текст түсін қатаң бекіту: Streamlit теңшелген теммен сәйкес келмесе де мәтін оқылсын */
.stApp,.stApp p,.stApp li,.stApp label,.stApp span,.stApp div[data-testid="stMarkdownContainer"],
[data-testid="stWidgetLabel"] *,[data-testid="stSidebar"] *,.stRadio label,.stSelectbox label,.stMultiSelect label{color:@text@}
.stApp h1,.stApp h2,.stApp h3,.stApp h4{color:@head@!important}
[data-testid="stCaptionContainer"],[data-testid="stCaptionContainer"] *{color:@soft@!important}
[data-testid="stMetricLabel"] *{color:@soft@!important}
[data-testid="stMetricValue"],[data-testid="stMetricValue"] *{color:@head@!important}
button[data-baseweb="tab"]{color:@soft@}
button[data-baseweb="tab"][aria-selected="true"]{color:#b3262e}
.hero .t{color:#fff!important;font-size:1.55rem;font-weight:800;line-height:1.1}
.hero .s{color:#e8c4c7!important;font-size:.95rem;margin-top:2px}
.hero .badge{color:#ffd9dc!important}
.login-hero p{color:@soft@!important;margin:0 0 6px}
.stButton>button *,.stDownloadButton>button *,[data-testid="stFormSubmitButton"]>button *{color:#fff!important}
[data-baseweb="select"]>div,[data-baseweb="input"],[data-baseweb="base-input"],.stTextInput input,.stNumberInput input,
.stDateInput input{background:@inp@!important;color:@text@!important;border-color:@bd@!important}
[data-baseweb="select"] *{color:@text@!important}
[data-baseweb="popover"] *,[data-baseweb="menu"] *{background:@inp@;color:@text@}
[data-testid="stAlert"] *{color:@text@}
[data-testid="stDateInput"] [data-baseweb="input"],[data-testid="stDateInput"] [data-baseweb="base-input"],
[data-testid="stDateInput"] input,[data-testid="stTextArea"] textarea,[data-baseweb="textarea"],
[data-testid="stNumberInput"] [data-baseweb="input"],[data-testid="stMultiSelect"] [data-baseweb="select"]>div,
[data-baseweb="tag"]{background:@inp@!important;color:@text@!important;-webkit-text-fill-color:@text@!important}
[data-testid="stDateInput"],[data-testid="stDateInput"] *{background-color:@inp@!important;color:@text@!important;
       -webkit-text-fill-color:@text@!important;border-color:@bd@!important}
[data-testid="stDateInput"] [data-testid="stWidgetLabel"],[data-testid="stDateInput"] [data-testid="stWidgetLabel"] *{background:transparent!important}
[data-testid="stDateInput"]{background:transparent!important}
[data-testid="stDateInput"] input::placeholder{color:@soft@!important;-webkit-text-fill-color:@soft@!important}
[data-testid="stDateInput"] svg,[data-baseweb="select"] svg{fill:@text@!important}
[data-testid="stCode"],[data-testid="stCode"] pre,[data-testid="stCode"] code,[data-testid="stCode"] span,
pre,pre code,pre span,.stCodeBlock *{background:@inp@!important;color:@text@!important}
[data-testid="stCode"] pre{border:1px solid @bd@}
[data-baseweb="calendar"],[data-baseweb="calendar"] *{background:@inp@!important;color:@text@!important}
[data-baseweb="calendar"] [aria-selected="true"],[data-baseweb="calendar"] [aria-selected="true"] *{background:#b3262e!important;color:#fff!important}
@nav@
@flip@
</style>
"""

THEMES = {
    "dark": dict(g1="rgba(143,29,36,.38)", g2="rgba(107,18,24,.30)", bg="#0c0c0d", side="#120a0b", bd="#3a1519",
                 card1="#1d0f11", card2="#120a0b", text="#f2eaea", head="#ffffff", soft="#c9b3b5", inp="#1b1213",
                 form1="#1b0e10", form2="#110a0b", shadow="rgba(0,0,0,.45)",
                 navsel="rgba(179,38,46,.24)", navtxt="#ff9aa0", navhov="rgba(255,255,255,.06)"),
    "light": dict(g1="rgba(179,38,46,.10)", g2="rgba(179,38,46,.07)", bg="#fbf7f7", side="#f3e8e9", bd="#e3c6c9",
                  card1="#ffffff", card2="#fbf1f2", text="#2b1a1c", head="#1c0f11", soft="#6e5558", inp="#ffffff",
                  form1="#ffffff", form2="#fbf3f3", shadow="rgba(120,30,36,.14)",
                  navsel="rgba(179,38,46,.11)", navtxt="#b3262e", navhov="rgba(0,0,0,.05)"),
}
THEME_LABELS = {"Dark": "dark", "Light": "light"}


def _svg(paths: str) -> str:
    import urllib.parse
    svg = ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="black" stroke-width="1.8" '
           'stroke-linecap="round" stroke-linejoin="round">' + paths + '</svg>')
    return 'url("data:image/svg+xml,' + urllib.parse.quote(svg) + '")'


_U = '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/>'
NAV_ICONS = {
    "nav_sec": [
        '<path d="M3 3v18h18"/><path d="M7 15v3M12 9v9M17 5v13"/>',                                    # қалу көрсеткіші
        _U + '<path d="m16 11 2 2 4-4"/>',                                                             # куратор
        _U + '<path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>',            # топтар
        '<path d="M6 9H4.5a2.5 2.5 0 0 1 0-5H6M18 9h1.5a2.5 2.5 0 0 0 0-5H18M4 22h16"/>'
        '<path d="M18 2H6v7a6 6 0 0 0 12 0V2zM10 14.7V17M14 14.7V17M8 22h8"/>',                        # рейтинг
        '<path d="m21.7 18-8-14a2 2 0 0 0-3.5 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.7-3z"/><path d="M12 9v4M12 17h.01"/>',
        '<path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/>',                               # тізімдер
        '<path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5z"/><path d="M14 2v6h6M16 13H8M16 17H8M10 9H8"/>',
        '<path d="M19 7V4a1 1 0 0 0-1-1H5a2 2 0 0 0 0 4h15a1 1 0 0 1 1 1v4h-3a2 2 0 0 0 0 4h3a1 1 0 0 0 1-1v-2"/>'
        '<path d="M3 5v14a2 2 0 0 0 2 2h15a1 1 0 0 0 1-1v-4"/>',                                       # суммалар
    ],
    "nav_area": [
        '<path d="M3 21h18M5 21V7l8-4v18M19 21V11l-6-4M9 9v.01M9 12v.01M9 15v.01M9 18v.01"/>',        # ішкі
        '<circle cx="12" cy="12" r="10"/><path d="M2 12h20"/>'
        '<path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/>',
    ],
}


def nav_css() -> str:
    out = []
    for key, icons in NAV_ICONS.items():
        g = f'[data-testid="stSidebar"] .st-key-{key}'
        out.append(f'{g} [role="radiogroup"]{{gap:4px;flex-direction:column;align-items:stretch}}')
        out.append(f'{g} [role="radiogroup"] label{{width:100%;margin:0!important;padding:11px 14px!important;border-radius:12px;'
                   f'border:1px solid transparent;cursor:pointer;display:flex;align-items:center;gap:12px;'
                   f'transition:background .15s,border-color .15s}}')
        out.append(f'{g} [role="radiogroup"] label>div:not(:has([data-testid="stMarkdownContainer"])){{display:none!important}}')
        out.append(f'{g} [role="radiogroup"] label p,{g} [role="radiogroup"] label span{{font-size:.98rem;font-weight:500}}')
        out.append(f'{g} [role="radiogroup"] label::before{{content:"";width:20px;height:20px;flex:none;background-color:currentColor;opacity:.8;'
                   f'-webkit-mask:var(--ic) center/contain no-repeat;mask:var(--ic) center/contain no-repeat}}')
        out.append(f'{g} [role="radiogroup"] label:hover{{background:@navhov@}}')
        out.append(f'{g} [role="radiogroup"] label:has(input:checked){{background:@navsel@;border-color:#b3262e}}')
        out.append(f'{g} [role="radiogroup"] label:has(input:checked) *{{color:@navtxt@!important;font-weight:700}}')
        for i, p in enumerate(icons, 1):
            out.append(f'{g} [role="radiogroup"] label:nth-of-type({i}){{--ic:{_svg(p)}}}')
    return "\n".join(out)


def native_theme() -> str:
    """Streamlit-тің өз темасы (кестелер мен графиктер соған бағынады). Анықтау мүмкін болмаса — config.toml бойынша қараңғы."""
    try:
        t = st.context.theme.type
        return t if t in ("light", "dark") else "dark"
    except Exception:
        return "dark"


def theme_css(mode: str) -> str:
    out = CSS.replace("@nav@", nav_css())
    for k, v in THEMES[mode].items():
        out = out.replace(f"@{k}@", v)
    # кесте мен графиктер Streamlit-тің өз темасымен салынады — таңдалған тема басқа болса, түсін айналдырамыз
    flip = ('[data-testid="stDataFrame"],[data-testid="stArrowVegaLiteChart"],[data-testid="stVegaLiteChart"]'
            '{filter:invert(.92) hue-rotate(180deg)}') if mode != native_theme() else ""
    return out.replace("@flip@", flip)


def hero(title: str, subtitle: str, badge: str = ""):
    b = f'<span class="badge">{badge}</span>' if badge else ""
    st.markdown(f'<div class="hero">{logo_img(46)}<div><div class="t">{title}</div>'
                f'<div class="s">{subtitle}</div></div>{b}</div>', unsafe_allow_html=True)

THEME = THEME_LABELS.get(st.session_state.get("theme_pick"), "dark")
st.markdown(theme_css(THEME), unsafe_allow_html=True)

# ------------------------------------------------------------------ кіру
if "role" not in st.session_state:
    if not ADMIN_PW and not STAFF_PW:
        st.session_state["role"] = "admin"                # құпиясөз қойылмаған — жергілікті сынақ
        st.session_state["open"] = True
    else:
        _, mid, _ = st.columns([1, 1.15, 1])
        with mid:
            st.markdown(f'<div class="login-hero"><div class="logo">{logo_img(78)}</div>'
                        '<h1>JUZ40 IELTS</h1><p>One team, one dream, no time to rest, IELTS team is the best</p></div>', unsafe_allow_html=True)
            with st.form("login"):
                pw = st.text_input("Құпиясөз", type="password", placeholder="Құпиясөзді енгізіңіз")
                go = st.form_submit_button("Кіру", use_container_width=True)
            if go:
                if ADMIN_PW and pw == ADMIN_PW:
                    st.session_state["role"] = "admin"; st.rerun()
                elif STAFF_PW and pw == STAFF_PW:
                    st.session_state["role"] = "staff"; st.rerun()
                else:
                    st.error("Құпиясөз дұрыс емес. Қайта көріңіз.")
            st.markdown('<p style="text-align:center;margin:14px 0 0;font-size:.9rem">Құпиясөзді админнен алыңыз. '
                        '<a href="https://t.me/a_erkebu1ann" target="_blank" rel="noopener" '
                        'style="color:#e5484d;font-weight:600;text-decoration:none">Админмен байланысу</a></p>',
                        unsafe_allow_html=True)
        st.stop()
ROLE = st.session_state["role"]
ADMIN = ROLE == "admin"
for _e in sync_remote(ROSTER_URL, PAYMENTS_URL):
    st.error(_e)
if not ROSTER.exists():
    st.error("Куратор тізімі табылмады. Админ ROSTER_URL (Google Sheets сілтемесі) қосуы немесе data/roster.xlsx жүктеуі керек.")
    st.stop()


def reports_gate(key: str) -> bool:
    """«Отчеттер» бөлімдеріне қосымша құпиясөз. Бір рет енгізсе, екі бөлім де ашылады."""
    if not REPORTS_PW:
        return True
    if st.session_state.get("reports_ok"):
        if st.button("🔒 Отчеттерді қайта жабу", key=f"lock_{key}"):
            st.session_state["reports_ok"] = False
            st.rerun()
        return True
    st.markdown('<div class="login-hero" style="margin:3vh 0 10px"><h1 style="font-size:1.6rem">🔒 Отчеттер жабық</h1>'
                '<p>Бұл бөлімді көру үшін қосымша құпиясөз керек</p></div>', unsafe_allow_html=True)
    _, mid, _ = st.columns([1, 1.2, 1])
    with mid:
        with st.form(f"gate_{key}"):
            pw = st.text_input("Отчеттер құпиясөзі", type="password", key=f"gpw_{key}", placeholder="Құпиясөзді енгізіңіз")
            go = st.form_submit_button("Ашу", use_container_width=True)
        if go:
            if pw == REPORTS_PW:
                st.session_state["reports_ok"] = True
                st.rerun()
            else:
                st.error("Құпиясөз дұрыс емес")
    return False


def load_settings() -> dict:
    try:
        return json.loads(SETTINGS.read_text(encoding="utf-8"))
    except Exception:
        return {}


def money(v) -> str:
    return f"{v:,.0f}".replace(",", " ") + " ₸"


# ------------------------------------------------------------------ деректер
@st.cache_data(ttl=300, show_spinner="Кесте Google-дан жүктелуде…")
def load_from_link(url: str):
    with tempfile.TemporaryDirectory() as d:
        p = core.fetch_google_sheet(url, Path(d) / "up.xlsx")
        return p.read_bytes(), datetime.now().strftime("%d.%m.%Y %H:%M")


@st.cache_data(show_spinner="Есептелуде…")
def compute(up_bytes: bytes, sheets: str, roster_stamp: float):
    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "up.xlsx"; p.write_bytes(up_bytes)
        up = core.load_up(p, sheet_filter=sheets)
        rep = core.build_report(core.load_roster(ROSTER), up)
        xlsx = core.write_report_xlsx(rep, Path(d) / "report.xlsx").read_bytes()
    return rep, list(up.sheets), xlsx


with st.sidebar:
    st.markdown(f"**{'👑 Админ' if ADMIN else '👤 Қызметкер'}**")
    if st.session_state.get("open"):
        st.warning("Құпиясөз қойылмаған — сайт ашық. .env-ке ADMIN_PASSWORD және STAFF_PASSWORD жазыңыз.")
    elif st.button("Шығу", key="btn_logout"):
        del st.session_state["role"]; st.rerun()
    if st.button("🔄 Деректерді жаңарту", key="btn_refresh"):
        load_from_link.clear(); compute.clear(); sync_remote.clear()
    st.radio("Theme", list(THEME_LABELS), horizontal=True, key="theme_pick")
    st.divider()
    AREA = "ІШКІ"
    if ADMIN:
        AREA = st.radio("Бет", ["JUZ40 IELTS ІШКІ", "JUZ40 IELTS СЫРТҚЫ"], key="nav_area").split()[-1]
    SEC_STAFF = ["Қалу көрсеткіші", "Куратор бойынша", "Топтар бойынша", "Рейтинг", "Мәселелер", "Тізімдер"]
    SEC_ADMIN = ["Отчеттер", "Суммалар"]
    SECTION = None
    if AREA == "ІШКІ":
        SECTION = st.radio("Бөлім", SEC_STAFF + (SEC_ADMIN if ADMIN else []), key="nav_sec")
    link, sheets, up_file = DEF_URL, DEF_SHEETS, None
    if ADMIN:
        st.divider()
        with st.expander("⚙️ Баптаулар", expanded=False):
            link = st.text_input("UP Google Sheets сілтемесі", DEF_URL,
                                 help="Кесте «сілтемесі бар кез келген адам — Көруші» болып бөлісілуі керек")
            sheets = st.text_input("Парақ атауында (ай)", DEF_SHEETS,
                                   help="Үтірмен: қазан, IELTS қыркүйек. Бос — барлық парақ.")
            up_file = st.file_uploader("…немесе UP Excel файлы", type=["xlsx"], key="upx")
            st.caption(f"Куратор тізімі: {ROSTER.name}")
            new = st.file_uploader("Куратор тізімін жаңарту (roster.xlsx)", type=["xlsx"], key="roster")
            if new is not None and st.button("Тізімді ауыстыру", key="btn_roster"):
                ROSTER.write_bytes(new.getvalue()); compute.clear(); st.success("Жаңартылды")

# ------------------------------------------------------------------ негізгі бет
ISSUE_PLAIN = {
    "Табылмады": ("Оқушы тізімде табылмады", "UP-тағы аты мен нөмірін тексеріңіз. Оқушы куратор тізімінде болмауы мүмкін."),
    "Нөмір сәйкес емес": ("Нөмірі басқа", "Оқушы аты бойынша табылды, төлемі есептелді. UP-тағы нөмірді платформадағыға сәйкестендіріңіз."),
    "Деңгей сәйкес емес": ("Деңгейі басқа", "Оқушы UP-та бір деңгей парағында, платформада басқа деңгейде тұр. Төлемі платформадағы тобында есептелді."),
    "Телефон форматы қате": ("Нөмірі дұрыс емес не жоқ", "UP-қа дұрыс нөмір жазыңыз: 87001234567 немесе +77001234567."),
    "Телефон сәйкес, аты басқа": ("Нөмірі бірдей, аты басқа", "Бұл басқа адам болуы мүмкін. Аты мен нөмірін екі жақта тексеріңіз."),
    "Аты бірдей бірнеше оқушы": ("Аты бірдей оқушылар бар", "Қайсысы екені анықталмады, төлемі есептелмеді. UP-қа дұрыс нөмірін жазыңыз."),
    "UP-та қайталанған": ("UP-та екі рет жазылған", "Бір рет есептелді. Артық қатарды өшіруге болады."),
    "Белгі түсініксіз": ("«иә» дұрыс жазылмаған", "Ұяшыққа дәл «иә» деп жазыңыз, сонда төлем есепке қосылады."),
    "Жасыл, бірақ белгі қарама-қайшы": ("Қатар жасыл, бірақ «жоқ» деп тұр", "Жасыл қатар төленді деп есептелді. Дұрыс болмаса, қатардың түсін алып тастаңыз."),
}


def curator_students(rep):
    """Куратор таңдаңыз → оның оқушылары: кім төледі, кім төлемеді."""
    st.subheader("👩‍🏫 Куратор бойынша оқушылар")
    st.caption("Кураторды таңдаңыз — оның барлық оқушылары және кім келесі айға төледі, кім төлемегені көрінеді.")
    stats = {}
    for (l, c, g), (t, p_, gg) in rep.by_group.items():
        a = stats.setdefault((l, c), [0, 0, 0]); a[0] += t; a[1] += p_; a[2] += gg
    if not stats:
        return
    keys = sorted(stats, key=lambda k: (k[0], k[1]))
    labels = {k: f"{k[1]} — {k[0]} ({stats[k][1]}/{stats[k][0]})" for k in keys}
    sel = st.selectbox("Куратор", keys, format_func=lambda k: labels[k], key="cur_pick")
    lv, cur = sel
    tot, paid_n, green_n = stats[sel]
    dates = {(x["Деңгей"], x["Куратор"], x["Студент"]): x["Күні"] for x in rep.payments}
    rows = []
    for r in rep.paid_students:                       # (деңгей, куратор, топ, аты, тел, түрі)
        if r[0] == lv and r[1] == cur:
            d = dates.get((r[0], r[1], r[3]))
            rows.append({"Оқушы": r[3], "Телефон": r[4], "Төлем": "🟩 Бірнеше айға" if "жасыл" in str(r[5]) else "✅ Төледі",
                         "Төлем күні": d.strftime("%d.%m") if d else "", "Топ": core.short_group(r[1], r[2])})
    for r in rep.unpaid_students:                     # (деңгей, куратор, топ, аты, тел)
        if r[0] == lv and r[1] == cur:
            rows.append({"Оқушы": r[3], "Телефон": r[4], "Төлем": "❌ Төлемеді", "Төлем күні": "",
                         "Топ": core.short_group(r[1], r[2])})
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Оқушы", tot)
    m2.metric("Төледі", f"{paid_n} ({paid_n / tot * 100:.0f}%)" if tot else 0)
    m3.metric("🟩 Бірнеше айға", green_n)
    m4.metric("Төлемеген", tot - paid_n)
    f = st.radio("Көрсету", ["Барлығы", "Төледі", "Төлемеді"], horizontal=True, key="cur_filter")
    df = pd.DataFrame(rows)
    if len(df):
        if f == "Төледі":
            df = df[~df["Төлем"].str.startswith("❌")]
        elif f == "Төлемеді":
            df = df[df["Төлем"].str.startswith("❌")]
        df = df.assign(_o=df["Төлем"].str.startswith("❌")).sort_values(["_o", "Оқушы"]).drop(columns="_o")
        st.dataframe(df.reset_index(drop=True), use_container_width=True, hide_index=True)
        st.download_button("⬇️ Осы куратордың тізімін жүктеу (CSV)", df.to_csv(index=False).encode("utf-8-sig"),
                           f"{cur}_{lv}.csv".replace(" ", "_").replace("|", ""), "text/csv", key="dl_cur")
    else:
        st.info("Бұл кураторда оқушы жоқ.")


def ichki_summary(rep, xlsx):
    tot = sum(v[0] for v in rep.by_group.values())
    paid = sum(v[1] for v in rep.by_group.values())
    green = sum(v[2] for v in rep.by_group.values())
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Барлық оқушы", tot)
    c2.metric("Келесі айға төледі", f"{paid} ({paid / tot * 100:.1f}%)" if tot else 0)
    c3.metric("🟩 Бірнеше айға алдын ала төледі", green)
    c4.metric("Бір айға төледі («иә»)", paid - green)
    st.download_button("⬇️ Толық Excel есепті жүктеу", xlsx, "report.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_report")


def sec_retention(rep, xlsx):
    ichki_summary(rep, xlsx)
    st.subheader("Деңгей бойынша: келесі айға қалған оқушылар")
    st.caption("Қалу көрсеткіші = келесі айға төлеген оқушылар ÷ деңгейдегі барлық оқушылар.")
    la = pd.DataFrame(core.level_analytics(rep))
    for col, (_, r) in zip(st.columns(len(la)), la.iterrows()):
        col.metric(r["Деңгей"], f"{r['Retention %']}%", f"{r['Төледі']} / {r['Оқушы']} оқушы төледі", delta_color="off")
        col.caption(f"🟩 бірнеше айға: {r['Бірнеше айға (🟩)']} · бір айға: {r['Тек 1 айға («иә»)']} · төлемеген: {r['Төлемеген']}")
    show = la[["Деңгей", "Оқушы", "Төледі", "Retention %", "Бірнеше айға (🟩)", "Тек 1 айға («иә»)", "Төлемеген"]].rename(
        columns={"Оқушы": "Оқушы саны", "Retention %": "Қалу көрсеткіші, %", "Бірнеше айға (🟩)": "Бірнеше айға",
                 "Тек 1 айға («иә»)": "Бір айға"})
    st.dataframe(show.set_index("Деңгей"), use_container_width=True)
    comp = la.melt(id_vars=["Деңгей"], value_vars=["Бірнеше айға (🟩)", "Тек 1 айға («иә»)", "Төлемеген"],
                   var_name="Санат", value_name="Саны")
    comp["Санат"] = comp["Санат"].replace({"Бірнеше айға (🟩)": "Бірнеше айға төледі",
                                           "Тек 1 айға («иә»)": "Бір айға төледі", "Төлемеген": "Төлемеген"})
    st.altair_chart(alt.Chart(comp).mark_bar().encode(
        y=alt.Y("Деңгей:N", title=None), x=alt.X("Саны:Q", stack="normalize", title="Үлесі"),
        color=alt.Color("Санат:N", scale=alt.Scale(domain=["Бірнеше айға төледі", "Бір айға төледі", "Төлемеген"],
                                                   range=["#a3202c", "#e58a90", "#4a4a50"]),
                        legend=alt.Legend(orient="bottom", title=None)),
        tooltip=["Деңгей", "Санат", "Саны"]).properties(height=max(120, 60 * len(la))), use_container_width=True)

    dl = pd.DataFrame(core.daily_by_level(rep))
    if len(dl):
        dl["Күні"] = pd.to_datetime(dl["Күні"])
        mode = st.radio("Графикте", ["Әр күнгі продажа", "Жиынтық қалу %"], horizontal=True, key="daily_mode")
        daily = mode.startswith("Әр күнгі")
        st.markdown("**Әр күні неше оқушы төледі**" if daily else "**Қалу көрсеткіші күн сайын қалай өсті**")
        st.caption("Тек төлем күні жазылған төлемдер." + (" Әр нүкте — сол күні жасалған продажа саны." if daily else ""))
        st.altair_chart(alt.Chart(dl).mark_line(point=True).encode(
            x=alt.X("Күні:T", title=None, axis=alt.Axis(format="%d.%m", tickCount="day", labelAngle=-90, labelOverlap=False, orient="bottom")),
            y=alt.Y("Жаңа төлем:Q", title="Күніне төлем саны") if daily else alt.Y("Retention %:Q", title="Қалу көрсеткіші, %"),
            color=alt.Color("Деңгей:N", scale=alt.Scale(range=SERIES), legend=alt.Legend(orient="bottom", title=None)),
            tooltip=["Күні", "Деңгей", "Жаңа төлем", "Жиынтық", "Retention %"]).properties(height=280),
            use_container_width=True)
        piv = dl.pivot(index="Күні", columns="Деңгей", values="Жаңа төлем")
        piv = piv[piv.sum(axis=1) > 0]
        piv.index = piv.index.strftime("%d.%m")
        st.markdown("**Күніне неше оқушы төледі**")
        st.dataframe(piv, use_container_width=True)
    nd = int(la["Күні жоқ төлем"].sum())
    if nd:
        st.caption(f"ℹ️ {nd} төлемнің күні UP-та жазылмаған (көбі жасыл қатарлар). Олар жалпы санға кіреді, "
                   "бірақ күн бойынша графикте көрінбейді.")


def sec_curators(rep):
    curator_students(rep)


def sec_groups(rep):
    st.subheader("Топтар бойынша есеп")
    rows = [{"Деңгей": l, "Куратор": c, "Топ": core.short_group(c, g), "Оқушы саны": t, "Төледі": p_,
             "Төлеген %": round(p_ / t * 100) if t else 0, "Бірнеше айға": gg, "Бір айға": p_ - gg}
            for (l, c, g), (t, p_, gg) in rep.by_group.items()]
    df = pd.DataFrame(rows).sort_values(["Деңгей", "Төлеген %"], ascending=[True, False])
    lv = st.multiselect("Деңгей таңдаңыз", sorted(df["Деңгей"].unique()), default=sorted(df["Деңгей"].unique()))
    st.dataframe(df[df["Деңгей"].isin(lv)], use_container_width=True, hide_index=True,
                 column_config={"Төлеген %": st.column_config.ProgressColumn("Төлеген %", min_value=0, max_value=100, format="%d%%")})


def sec_rating(rep):
    st.subheader("🏆 Рейтинг")
    dated = [x["Күні"] for x in rep.payments if x["Күні"]]
    mode = st.radio("Қай кезең", ["Барлығы (жасыл қатарларды қоса)", "Таңдалған күндер (төлем күні жазылғандар ғана)"], horizontal=True)
    start = end = None
    if mode.startswith("Таңдалған"):
        if not dated:
            st.warning("UP-та төлем күні жазылған төлем жоқ."); return
        lo, hi = min(dated), max(dated)
        pr = st.date_input("Күндер аралығы", (lo, hi), min_value=lo, max_value=hi)
        start, end = (pr if len(pr) == 2 else (pr[0], pr[0]))
        st.caption("Тек осы күндерде төлем жасағандар есептеледі.")
    rc1, rc2, rc3 = st.columns([2, 1, 1])
    what = rc1.radio("Нені салыстырамыз", ["Кураторлар", "Топтар"], horizontal=True)
    n = rc2.slider("Қанша көрсету", 3, 50, 5)
    order = rc3.radio("Тәртібі", ["Үздіктерден бастап", "Төменнен бастап"], horizontal=True)
    data = core.rank_curators(rep, start, end) if what == "Кураторлар" else core.rank_groups(rep, start, end)
    key = "Куратор" if what == "Кураторлар" else "Топ"
    data = data[:n] if order.startswith("Үздік") else data[::-1][:n]
    rdf = pd.DataFrame(data)
    rdf["Белгі"] = rdf["№"].astype(str) + ". " + rdf[key]
    st.altair_chart(alt.Chart(rdf).mark_bar().encode(
        x=alt.X("%:Q", title="Төлем %", scale=alt.Scale(domain=[0, max(100 if not start else 10, float(rdf["%"].max()))])),
        y=alt.Y("Белгі:N", sort=None, title=None),
        color=alt.Color("%:Q", scale=alt.Scale(scheme="reds"), legend=None),
        tooltip=[key, "Оқушы", "Төледі", "🟩", "%"]).properties(height=max(120, 38 * len(rdf))), use_container_width=True)
    st.dataframe(rdf.drop(columns=["Белгі"]).set_index("№"), use_container_width=True)
    st.caption("Рейтинг төлем пайызы бойынша. Пайыз бірдей болса, көбірек оқушы төлеген жоғары тұрады. "
               "Оқушысы аз топта пайыз тұрақсыз болуы мүмкін.")


def sec_issues(rep, admin):
    st.subheader(f"⚠️ Мәселелер ({len(rep.issues)})")
    st.caption("UP кестесінде түзету керек болуы мүмкін жерлер. Төлем көбінде есептелген, бірақ дұрыстап қойсаңыз, есеп анық болады.")
    if rep.issues:
        rows = []
        for i in rep.issues:
            what_, todo = ISSUE_PLAIN.get(i["Түрі"], (i["Түрі"], ""))
            r = {"Не болды": what_, "Не істеу керек": todo, "UP-тағы оқушы": i["UP аты"], "UP қатар": i["UP қатар"],
                 "UP нөмірі": i["UP телефон"], "Куратор / топ": i["Куратор / топ"]}
            if admin:
                r["Толық түсініктеме"] = i["Түсініктеме"]
            rows.append(r)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    else:
        st.success("Бәрі дұрыс, тексеретін ештеңе жоқ")


def sec_lists(rep):
    st.subheader("📋 Төлегендер және төлемегендер")
    base = ["Деңгей", "Куратор", "Топ", "Студент", "Телефон"]
    t2, t3 = st.tabs(["Төлегендер", "Төлемегендер"])
    with t2:
        pdf = pd.DataFrame(rep.payments)[base + ["Күні", "Түрі"]]
        st.dataframe(pdf, use_container_width=True, hide_index=True)
    with t3:
        st.dataframe(pd.DataFrame(rep.unpaid_students, columns=base), use_container_width=True, hide_index=True)


# ------------------------------------------------------------------ админ: отчеттер
def reports_page(rep):
    if not reports_gate("ichki"):
        return
    st.subheader("📝 Отчеттер")
    st.caption("Топтар мен форматты reports.py файлынан өзгертуге болады.")
    dated = [x["Күні"] for x in rep.payments if x["Күні"]]
    c1, c2 = st.columns(2)
    day = c1.date_input("Отчет күні", max(dated) if dated else date.today(), format="DD.MM.YYYY")
    start = c2.date_input("Жалпы отчет басталатын күн", date(day.year, 9, 17), format="DD.MM.YYYY")
    for w in reports.report_warnings(rep):
        st.warning(w)
    daily = reports.daily_report(rep, day)
    t_ishki, t_ltl = reports.total_reports(rep, start, day)
    for title, text, fn in [("Күндік отчет", daily, f"kundik_{day:%d.%m}.txt"),
                            ("ІШКІ САТУ — жалпы отчет", t_ishki, f"ishki_jalpy_{day:%d.%m}.txt"),
                            ("LEVEL TO LEVEL — жалпы отчет", t_ltl, f"ltl_jalpy_{day:%d.%m}.txt")]:
        st.markdown(f"**{title}**")
        st.code(text, language=None)          # оң жақ жоғарыдағы түйме — көшіру
        st.download_button("⬇️ .txt жүктеу", text, fn, key=fn)
    st.markdown("**Үшеуі бірге** (бір рет көшіру үшін)")
    st.code("\n\n".join([daily, t_ishki, t_ltl]), language=None)


# ------------------------------------------------------------------ админ: суммалар
def money_page(rep):
    st.subheader("💰 Суммалар")
    st.caption("«Төленген сумма» — UP-тағы «Төлем суммасы» бағанынан (тек ол толтырылған қатарлар). "
               "«Баға бойынша» — UP-тағы «СУММА» бағанынан (пакет бағасы), барлық төлеген оқушы үшін. "
               "Жасыл қатарларда нақты сома жазылмайды, сондықтан соманы «Баға бойынша» бағанынан қараңыз.")
    p = pd.DataFrame(rep.payments)
    if p.empty:
        st.info("Төлем жоқ"); return
    g = p.groupby("Деңгей").agg(Төлеген_оқушы=("Студент", "count"), Төленген=("Төленген сумма", "sum"),
                                Баға=("Баға (СУММА)", "sum")).reset_index()
    g.columns = ["Деңгей", "Төлеген оқушы", "Төленген сумма", "Баға бойынша"]
    t = pd.DataFrame([{"Деңгей": "ЖАЛПЫ", "Төлеген оқушы": g["Төлеген оқушы"].sum(),
                       "Төленген сумма": g["Төленген сумма"].sum(), "Баға бойынша": g["Баға бойынша"].sum()}])
    g = pd.concat([g, t])
    m1, m2 = st.columns(2)
    m1.metric("Төленген сумма (UP-та жазылған)", money(t["Төленген сумма"].iloc[0]))
    m2.metric("Баға бойынша (СУММА бағаны)", money(t["Баға бойынша"].iloc[0]))
    for c in ("Төленген сумма", "Баға бойынша"):
        g[c] = g[c].map(money)
    st.dataframe(g.set_index("Деңгей"), use_container_width=True)
    c = p.groupby("Куратор").agg(Оқушы=("Студент", "count"), Баға=("Баға (СУММА)", "sum"),
                                 Төленген=("Төленген сумма", "sum")).reset_index().sort_values("Баға", ascending=False)
    st.markdown("**Кураторлар бойынша**")
    c["Баға"] = c["Баға"].map(money); c["Төленген"] = c["Төленген"].map(money)
    c.columns = ["Куратор", "Төлеген оқушы", "Баға бойынша", "Төленген сумма"]
    st.dataframe(c, use_container_width=True, hide_index=True)
    d = p[p["Күні"].notna()].groupby("Күні").agg(Оқушы=("Студент", "count"), Төленген=("Төленген сумма", "sum"),
                                                  Баға=("Баға (СУММА)", "sum")).reset_index()
    if len(d):
        st.markdown("**Күн бойынша** (төлем күні жазылғандар)")
        d["Күні"] = d["Күні"].map(lambda x: x.strftime("%d.%m")); d["Төленген"] = d["Төленген"].map(money)
        d["Баға"] = d["Баға"].map(money)
        d.columns = ["Күні", "Төлеген оқушы", "Төленген сумма", "Баға бойынша"]
        st.dataframe(d, use_container_width=True, hide_index=True)



def render_ichki():
    """JUZ40 IELTS ІШКІ: бар оқушылардың келесі айға төлемі (UP кестесі бойынша)."""
    try:
        if up_file is not None:
            up_bytes, stamp = up_file.getvalue(), "жүктелген файл"
        elif link:
            up_bytes, stamp = load_from_link(link)
            stamp = f"соңғы жаңарту {stamp}"
        else:
            st.info("Кесте әлі жалғанбаған. Админ UP кестесінің сілтемесін қосуы керек." if not ADMIN else
                    "Сол жақтағы «⚙️ Баптаулар» бөліміне UP кестесінің Google Sheets сілтемесін қойыңыз немесе Excel жүктеңіз.")
            return
        rep, used_sheets, xlsx = compute(up_bytes, sheets, ROSTER.stat().st_mtime)
    except core.UpFileError as e:
        st.error(str(e)); return
    st.caption(f"Дерек көзі: {stamp} · оқылған парақтар: {', '.join(used_sheets)}")
    if SECTION == SEC_STAFF[0]:
        sec_retention(rep, xlsx)
    elif SECTION == SEC_STAFF[1]:
        sec_curators(rep)
    elif SECTION == SEC_STAFF[2]:
        sec_groups(rep)
    elif SECTION == SEC_STAFF[3]:
        sec_rating(rep)
    elif SECTION == SEC_STAFF[4]:
        sec_issues(rep, ADMIN)
    elif SECTION == SEC_STAFF[5]:
        sec_lists(rep)
    elif ADMIN and SECTION == SEC_ADMIN[0]:
        reports_page(rep)
    elif ADMIN and SECTION == SEC_ADMIN[1]:
        money_page(rep)


# ------------------------------------------------------------------ СЫРТҚЫ (тек админ)
@st.cache_data(show_spinner="Төлемдер оқылуда…")
def load_external(stamp: float, roster_stamp: float):
    pays = ext.load_payments(PAY_FILE)
    ext.mark_known(pays, core.load_roster(ROSTER))
    return pays


def external_page():
    """JUZ40 IELTS СЫРТҚЫ: платформада келесі айға (Қараша) төлем жасаған оқушылар тізімі, анализ, рейтинг."""
    with st.expander("📥 Төлемдер файлы (платформадан экспорт)", expanded=not PAY_FILE.exists()):
        up = st.file_uploader("Төлемдер Excel файлын жүктеңіз («Төлемдер» парағы)", type=["xlsx"], key="payx")
        if up is not None and st.button("Файлды сақтау және жаңарту", key="btn_payfile"):
            PAY_FILE.parent.mkdir(exist_ok=True)
            PAY_FILE.write_bytes(up.getvalue()); load_external.clear(); st.success("Сақталды")
        if PAY_FILE.exists():
            st.caption("Қазіргі файл жаңартылған: " + datetime.fromtimestamp(PAY_FILE.stat().st_mtime).strftime("%d.%m.%Y %H:%M"))
    if not PAY_FILE.exists():
        st.info("Алдымен төлемдер файлын жүктеңіз."); return
    try:
        pays = load_external(PAY_FILE.stat().st_mtime, ROSTER.stat().st_mtime)
    except core.UpFileError as e:
        st.error(str(e)); return
    if not pays:
        st.warning(f"Файлда «{ext.NEXT_MONTH}» курсына төленген төлем табылмады."); return

    levels = sorted({p.level for p in pays}); mgrs = sorted({p.manager for p in pays})
    f1, f2, f3 = st.columns(3)
    lv = f1.multiselect("Деңгей", levels, default=levels)
    mg = f2.multiselect("Менеджер", mgrs, default=mgrs)
    dts = [p.date for p in pays if p.date]
    if dts:
        pr = f3.date_input("Күндер аралығы", (min(dts), max(dts)), min_value=min(dts), max_value=max(dts), format="DD.MM.YYYY")
        d0, d1 = (pr if len(pr) == 2 else (pr[0], pr[0]))
    else:
        d0 = d1 = None
    sel = [p for p in pays if p.level in lv and p.manager in mg and (d0 is None or p.date is None or d0 <= p.date <= d1)]
    full_all = [p for p in pays if not p.prepay]        # отчеттер сүзгіге тәуелсіз, предоплатасыз
    pre = [p for p in sel if p.prepay]          # предоплаталар негізгі көрсеткіштерге кірмейді
    sel = [p for p in sel if not p.prepay]
    s_ = ext.summary(sel)
    k = st.columns(5)
    k[0].metric("Келесі айға төлеген оқушы", s_["Оқушы (бірегей)"])
    k[1].metric("Жалпы сома", money(s_["Жалпы сома"]))
    k[2].metric("1 айға", s_["1 айға"])
    k[3].metric("2+ айға", s_["2+ айға"])
    k[4].metric("Куратор базасында бар", s_["Базада бар"])
    st.caption("«Қараша» курсына платформада толық төлем жасағандар. Предоплаталар бұл санға кірмейді, олар «💳 Предоплата» қойындысында. "
               "«Куратор базасында бар» — бұрыннан бар оқушы (телефон/аты бойынша), қалғандары — жаңа оқушылар.")

    t1, t2, t3, t4, t5 = st.tabs(["📋 Тізім", "📈 Анализ", "🏆 Рейтинг", f"💳 Предоплата ({len(pre)})", "📝 Отчеттер"])
    with t1:
        tdf = pd.DataFrame(ext.students_table(sel))
        st.dataframe(tdf, use_container_width=True, hide_index=True,
                     column_config={"Сома": st.column_config.NumberColumn("Сома, ₸", format="%d")})
        st.download_button("⬇️ Тізімді CSV жүктеу", tdf.to_csv(index=False).encode("utf-8-sig"), "syrtqy_qarasha.csv", "text/csv", key="dl_ext_csv")
    with t2:
        lvdf = pd.DataFrame(ext.by_key(sel, lambda p: p.level, "Деңгей")).drop(columns=["№"])
        st.markdown("**Деңгей бойынша**")
        st.dataframe(lvdf.set_index("Деңгей"), use_container_width=True)
        st.altair_chart(alt.Chart(lvdf).mark_bar(color="#b3262e").encode(
            x=alt.X("Деңгей:N", sort="-y", title=None), y=alt.Y("Оқушы:Q"), tooltip=["Деңгей", "Оқушы", "Сома"]).properties(height=240),
            use_container_width=True)
        dd = pd.DataFrame(ext.daily(sel))
        if len(dd):
            dd["Күні"] = pd.to_datetime(dd["Күні"])
            st.markdown("**Күн бойынша**")
            st.altair_chart(alt.Chart(dd).mark_bar(color="#e5484d").encode(
                x=alt.X("Күні:T", title=None), y=alt.Y("Жаңа төлем:Q", title="Күніне төлеген оқушы"),
                tooltip=["Күні", "Жаңа төлем", "Жиынтық", "Сома"]).properties(height=240), use_container_width=True)
            st.altair_chart(alt.Chart(dd).mark_line(point=True).encode(
                x=alt.X("Күні:T", title=None), y=alt.Y("Жиынтық:Q", title="Жиынтық оқушы саны"),
                tooltip=["Күні", "Жиынтық", "Жиынтық сома"]).properties(height=240), use_container_width=True)
            t = dd.copy(); t["Күні"] = t["Күні"].dt.strftime("%d.%m")
            st.dataframe(t.set_index("Күні"), use_container_width=True)
        mdf = pd.DataFrame(ext.by_key(sel, lambda p: f"{p.months} ай", "Неше айға")).drop(columns=["№"])
        st.markdown("**Неше айға төлегендер**")
        st.dataframe(mdf.set_index("Неше айға"), use_container_width=True)
    with t3:
        what = st.radio("Нені салыстырамыз", ["Менеджерлер", "Деңгейлер", "Қалалар"], horizontal=True)
        by = st.radio("Қандай көрсеткіш бойынша", ["Оқушы саны", "Сома"], horizontal=True)
        n = st.slider("Қанша көрсету", 3, 30, 10, key="extn")
        keyfn, label = {"Менеджерлер": (lambda p: p.manager, "Менеджер"), "Деңгейлер": (lambda p: p.level, "Деңгей"),
                        "Қалалар": (lambda p: p.city or "Көрсетілмеген", "Қала")}[what]
        rk = ext.by_key(sel, keyfn, label)
        rk.sort(key=lambda x: (-(x["Оқушы"] if by == "Оқушы саны" else x["Сома"]), str(x[label])))
        for i, x in enumerate(rk, 1):
            x["№"] = i
        rdf = pd.DataFrame(rk[:n])
        rdf["Белгі"] = rdf["№"].astype(str) + ". " + rdf[label]
        st.altair_chart(alt.Chart(rdf).mark_bar().encode(
            x=alt.X(f"{'Оқушы' if by == 'Оқушы саны' else 'Сома'}:Q", title=by),
            y=alt.Y("Белгі:N", sort=None, title=None),
            color=alt.Color(f"{'Оқушы' if by == 'Оқушы саны' else 'Сома'}:Q", scale=alt.Scale(scheme="reds"), legend=None),
            tooltip=[label, "Оқушы", "Сома", "1 айға", "2+ айға"]).properties(height=max(120, 38 * len(rdf))), use_container_width=True)
        show = rdf.drop(columns=["Белгі"]).set_index("№")
        show["Сома"] = show["Сома"].map(money)
        st.dataframe(show, use_container_width=True)
    with t4:
        st.caption("Алдын ала төлем (предоплата) жасаған оқушылар. Олар жоғарыдағы негізгі көрсеткіштерге, анализге және рейтингке кірмейді.")
        if not pre:
            st.info("Таңдалған сүзгі бойынша предоплата жоқ.")
        else:
            m1, m2 = st.columns(2)
            m1.metric("Предоплата саны", len(pre))
            m2.metric("Предоплата сомасы", money(sum(p.amount for p in pre)))
            pm = pd.DataFrame(ext.by_key(pre, lambda p: p.manager, "Менеджер"))[["№", "Менеджер", "Оқушы", "Сома"]]
            pm = pm.rename(columns={"Оқушы": "Предоплата саны"})
            st.markdown("**Менеджер бойынша**")
            pm["Сома"] = pm["Сома"].map(money)
            st.dataframe(pm.set_index("№"), use_container_width=True)
            st.markdown("**Предоплата тізімі**")
            pdf = pd.DataFrame(ext.students_table(pre))[["Күні", "Оқушы", "Телефон", "Деңгей", "Сома", "Менеджер"]]
            st.dataframe(pdf, use_container_width=True, hide_index=True)
    with t5:
        if not reports_gate("ext"):
            return
        st.caption("Отчеттер барлық менеджер мен деңгейді қамтиды (жоғарыдағы сүзгі әсер етпейді), предоплаталарсыз. "
                   "GE = BEGINNER + ELEMENTARY + PRE-INTER. Конверсия әзірге әрқашан 0%.")
        sett = load_settings()
        dts_all = [p.date for p in full_all if p.date]
        if not dts_all:
            st.info("Күні бар төлем жоқ."); return
        c1, c2 = st.columns(2)
        day = c1.date_input("Отчет күні", max(dts_all), format="DD.MM.YYYY", key="ext_day")
        start = c2.date_input("Жалпы отчет басталатын күн", min(dts_all), format="DD.MM.YYYY", key="ext_start")
        p1, p2, p3 = st.columns([1, 1, 1])
        plan_i = p1.number_input("IELTS жоспар (оқушы)", 0, 100000, int(sett.get("plan_ielts", 0)), key="plan_i",
                                 help="«38/220» деген жердегі 220. 0 болса, тек төлеген сан көрсетіледі.")
        plan_g = p2.number_input("GE жоспар (оқушы)", 0, 100000, int(sett.get("plan_ge", 0)), key="plan_g")
        if p3.button("💾 Жоспарды сақтау", key="btn_plan"):
            SETTINGS.parent.mkdir(exist_ok=True)
            SETTINGS.write_text(json.dumps({"plan_ielts": int(plan_i), "plan_ge": int(plan_g)}), encoding="utf-8")
            st.success("Сақталды")
        daily_t = reports.external_daily_report(full_all, day)
        total_t = reports.external_total_report(full_all, start, day, int(plan_i), int(plan_g), ext.NEXT_MONTH.upper())
        for title, text, fn in [("Күндік отчет", daily_t, f"syrtqy_kundik_{day:%d.%m}.txt"),
                                ("Жалпы отчет", total_t, f"syrtqy_jalpy_{day:%d.%m}.txt")]:
            st.markdown(f"**{title}**")
            st.code(text, language=None)
            st.download_button("⬇️ .txt жүктеу", text, fn, key="dl_" + fn)
        st.markdown("**Екеуі бірге** (бір рет көшіру үшін)")
        st.code(daily_t + "\n\n" + total_t, language=None)


# ------------------------------------------------------------------ бет
if ADMIN:
    hero("JUZ40 IELTS", "Келесі айға төлем есебі", "👑 Админ")
    if AREA == "ІШКІ":
        render_ichki()
    else:
        st.subheader("JUZ40 IELTS СЫРТҚЫ · келесі айға (Қараша) төлем жасағандар")
        external_page()
else:
    hero("JUZ40 IELTS ІШКІ", "Келесі айға төлем есебі", "👤 Қызметкер")
    render_ichki()
