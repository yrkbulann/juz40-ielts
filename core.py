"""Куратор топтары бойынша 'келесі айға төлем жасады' есебінің ядросы (Telegram-сыз, тестіленеді).

Сәйкестендіру: алдымен телефон (кез келген формат), болмаса аты (жазылу айырмашылығы ескеріледі).
Барлық күмәнді жағдайлар (басқа нөмір, басқа деңгей, қате формат, табылмаған) 'Мәселелер' тізіміне түседі.
"""
from __future__ import annotations
import re
import difflib
import datetime as dt
import ssl
import urllib.request
from collections import defaultdict, Counter
from dataclasses import dataclass, field
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ---- UP кестесіндегі бағандарды атауы бойынша іздеу (кіші әріппен) ----
NAME_KEYS = ("студент", "оқушы", "аты", "фио", "name", "имя")
PHONE_KEYS = ("телефон", "phone", "тел", "номер", "whatsapp", "ватсап")
LEVEL_KEYS = ("деңгей", "level", "уровень")
NEXT_KEYS = ("келесі ай", "следующ")          # «Келесі айға қалады ма?» бағаны
YES = {"иә"}                                    # менеджер төлегенін тек «иә» деп белгілейді
ALMOST_YES = {"иа", "иэ", "ия", "да", "yes", "+", "1", "ok", "ок", "иә."}   # «иә» деп жазғысы келген болуы мүмкін
PAID_KEYS = ("келесі ай", "келесі айға", "төлем", "төлед", "оплат", "paid", "төленді", "следующ")
PAID_TRUE = {"+", "✓", "✔", "✅", "1", "true", "иә", "да", "yes", "y", "x", "төледі", "төленді",
             "оплатил", "оплачено", "paid", "ок", "ok"}

_LAT2 = [("shch", "щ"), ("sh", "ш"), ("ch", "ч"), ("zh", "ж"), ("kh", "х"), ("ya", "я"), ("yu", "ю"),
         ("yo", "е"), ("ts", "ц")]
_LAT1 = {"a": "а", "b": "б", "c": "к", "d": "д", "e": "е", "f": "ф", "g": "г", "h": "х", "i": "и", "j": "ж",
         "k": "к", "l": "л", "m": "м", "n": "н", "o": "о", "p": "п", "q": "к", "r": "р", "s": "с", "t": "т",
         "u": "у", "v": "в", "w": "в", "x": "кс", "y": "и", "z": "з"}
_CYR = {"ә": "а", "ғ": "г", "қ": "к", "ң": "н", "ө": "о", "ү": "у", "ұ": "у", "һ": "х", "і": "и",
        "ы": "и", "ё": "е", "й": "и", "э": "е", "ъ": "", "ь": ""}


def fold_name(v) -> str:
    """Аттың жазылу нұсқаларын біріктіру: латын/кирилл, қазақ әріптері, сөз реті."""
    s = str(v or "").lower()
    for a, b in _LAT2:
        s = s.replace(a, b)
    s = "".join(_LAT1.get(ch, ch) for ch in s)
    s = "".join(_CYR.get(ch, ch) for ch in s)
    words = re.findall(r"[а-яa-z]+", s)
    return " ".join(sorted(words))


norm_name = fold_name   # кері үйлесімділік


def parse_phones(v):
    """Кез келген форматтан телефондарды алады: +7, 8, пробел, жақша, сызықша, бірнеше нөмір."""
    good, bad = [], []
    if v is None or v == "":
        return good, bad
    if isinstance(v, float) and v == int(v):
        v = int(v)
    for part in re.split(r"[;,/\n|]| или | и ", str(v)):
        d = re.sub(r"\D", "", part)
        if not d:
            continue
        if len(d) == 11 and d[0] in "78":
            d = d[1:]
        if len(d) == 10:
            good.append(d)
        else:
            bad.append(part.strip())
    return good, bad


def norm_phone(v) -> str:
    g, _ = parse_phones(v)
    return g[0] if g else ""


def fmt_phone(d: str) -> str:
    return f"+7{d}" if d else ""


def canon_level(text) -> str:
    t = str(text or "").upper()
    if not t.strip():
        return ""
    if "IELTS" in t:
        return "IELTS"
    if "PRE" in t or "B1" in t:
        return "PRE"
    if "ELEM" in t or "A2" in t:
        return "ELEM"
    if "BEG" in t or "A0" in t or "A1" in t:
        return "BEGIN"
    if "INTER" in t or "B2" in t:
        return "IELTS"
    return ""


def is_paid(v) -> bool:
    if v is True:
        return True
    if v is None or v is False:
        return False
    if isinstance(v, (int, float)):
        return v == 1
    return str(v).strip().lower() in PAID_TRUE


# ------------------------------- куратор тізімі -------------------------------
@dataclass
class Student:
    level: str
    curator: str
    name: str
    phone: str
    key_phone: str = ""
    key_name: str = ""
    group: str = ""


def load_roster(path) -> list[Student]:
    ws = openpyxl.load_workbook(path, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    head = [str(h or "").strip().lower() for h in rows[0]]

    def col(*names):
        for i, h in enumerate(head):
            if any(n in h for n in names):
                return i
        raise ValueError(f"Роста бағаны табылмады: {names}")
    ci, li, ni, pi = col("куратор"), col("деңгей", "level"), col("студент", "аты"), col("телефон")
    gi = next((i for i, h in enumerate(head) if h in ("топ", "group")), None)
    out = []
    for r in rows[1:]:
        if not r or not r[ni]:
            continue
        out.append(Student(str(r[li] or "—").strip(), str(r[ci] or "—").strip(), str(r[ni]).strip(),
                           str(r[pi] or ""), norm_phone(r[pi]), fold_name(r[ni]),
                           str(r[gi]).strip() if gi is not None and r[gi] else ""))
    return out


# ------------------------------------ UP ------------------------------------
@dataclass
class UpRow:
    row: int
    sheet: str
    name: str
    phones: list
    bad_phones: list
    raw_phone: str
    level: str
    green: bool = False      # қатар жасылмен боялған = бірнеше айға төленген
    months: int = 0          # «НЕШЕ АЙҒА» бағанынан (2 ай, 3 ай…)
    marker: str = ""         # «Келесі айға қалады ма?» ұяшығындағы мән
    pay_date: object = None  # «Төлем жасаған уақыты» (datetime.date) — көрсетілмесе None
    amount: float = 0.0      # «Төлем суммасы» (нақты төленген сома)
    plan: float = 0.0        # «Сумма / СУММА» бағаны (пакет бағасы)


@dataclass
class UpData:
    paid_rows: list = field(default_factory=list)
    paid_column: str = ""
    level_column: str = ""
    total_rows: int = 0
    sheets: list = field(default_factory=list)
    doubtful: list = field(default_factory=list)   # «иә» емес, бірақ соған ұқсас мәндер

    @property
    def paid_count(self):
        return len(self.paid_rows)

    @property
    def green_count(self):
        return sum(1 for r in self.paid_rows if r.green)


class UpFileError(Exception):
    pass


DATE_KEYS = ("төлем жасаған", "оплат")           # «Төлем жасаған уақыты» бағаны
DATE_RE = re.compile(r"(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})")


def parse_date(v):
    """Ұяшықтағы күнді date-қа айналдырады (datetime, '19.10.2026', '2026-10-19'); болмаса None."""
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        t = v.strip()
        m = DATE_RE.search(t)
        if m:
            d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
            y += 2000 if y < 100 else 0
            try:
                return dt.date(y, mo, d)
            except ValueError:
                return None
        m = re.match(r"(\d{4})-(\d{2})-(\d{2})", t)
        if m:
            try:
                return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                return None
    return None


def parse_num(v) -> float:
    if isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        t = re.sub(r"[\s\u00a0]", "", v).replace(",", ".")
        try:
            return float(t)
        except ValueError:
            return 0.0
    return 0.0


def fetch_google_sheet(url: str, dest) -> Path:
    """Google Sheets сілтемесінен кестені .xlsx (түстерімен) етіп жүктейді.
    Кесте «сілтемесі бар кез келген адам — көруші» болып бөлісілген болуы керек."""
    m = re.search(r"/spreadsheets/d/([\w-]+)", url or "")
    if not m:
        raise UpFileError("Бұл Google Sheets сілтемесі емес (…/spreadsheets/d/…/ болуы керек).")
    exp = f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=xlsx"
    try:
        req = urllib.request.Request(exp, headers={"User-Agent": "Mozilla/5.0"})
        try:                                   # macOS-та Python жүйелік сертификаттарды көрмеуі мүмкін
            import certifi
            ctx = ssl.create_default_context(cafile=certifi.where())
        except ImportError:
            ctx = ssl.create_default_context()
        with urllib.request.urlopen(req, timeout=60, context=ctx) as r:
            data = r.read()
    except Exception as e:  # noqa
        raise UpFileError(f"Кестені жүктей алмадым ({e}). Сілтеме бойынша қолжетімділік «көруші» болып ашық па?")
    if data[:2] != b"PK":
        raise UpFileError("Google кестені бермеді (жабық болуы мүмкін). Бөлісу: «Сілтемесі бар кез келген адам — Көруші».")
    Path(dest).write_bytes(data)
    return Path(dest)


MONTHS_RE = re.compile(r"(\d+)\s*ай", re.I)
NEGATIVE = {"жоқ", "нет", "no", "жок", "үлгермеу"}


def _is_green(cell) -> bool:
    """Ұяшық жасыл түспен боялған ба (Google Sheets жасылдары: B6D7A8, 93C47D, 00FF00…)."""
    f = cell.fill
    if f is None or f.fill_type != "solid" or f.fgColor is None or f.fgColor.type != "rgb":
        return False
    h = str(f.fgColor.rgb or "")
    if len(h) != 8:
        return False
    try:
        r, g, b = int(h[2:4], 16), int(h[4:6], 16), int(h[6:8], 16)
    except ValueError:
        return False
    return g >= r + 20 and g >= b + 20 and g >= 150 and r < 235      # анық жасыл реңк


def _green_row(cells, width=12) -> bool:
    """Қатардың басым бөлігі жасыл болса — қатар боялған деп есептеледі (жалғыз боялған ұяшық саналмайды)."""
    head = cells[:width]
    return len(head) >= 3 and sum(_is_green(c) for c in head) >= len(head) // 2


def _find_header(rows, paid_column):
    """Тақырып қатарын табады: (hi, ni, pis, qi, li, heads) немесе None."""
    for hi, r in enumerate(rows[:15]):
        heads = [str(h or "").strip() for h in r]
        low = [h.lower() for h in heads]
        ni = next((i for i, h in enumerate(low) if any(k in h for k in NAME_KEYS)), None)
        pis = [i for i, h in enumerate(low) if any(k in h for k in PHONE_KEYS) and i != ni]
        if paid_column:
            qi = next((i for i, h in enumerate(low) if h == paid_column.strip().lower()), None)
        else:
            qi = next((i for i, h in enumerate(low) if any(k in h for k in NEXT_KEYS)), None)
            if qi is None:
                c = [i for i, h in enumerate(low) if any(k in h for k in PAID_KEYS) and i != ni and i not in pis]
                qi = c[-1] if c else None
        if qi is not None and (ni is not None or pis):
            li = next((i for i, h in enumerate(low) if any(k in h for k in LEVEL_KEYS)
                       and i not in (ni, qi, *pis)), None)
            return hi, ni, pis, qi, li, heads
    return None


def load_up(path, paid_column: str = "", sheet_filter: str = "") -> UpData:
    """UP кестесінің ӘР парағын оқиды. Төлем белгісі: «Келесі айға қалады ма?» бағанында «иә»
    (баған не оның оң жақ көршісінде тұрса да болады — кей қатарлар бір баған жылжыған)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    up, first = UpData(), True
    flt = [x.strip().lower() for x in (sheet_filter or "").split(",") if x.strip()]
    for ws in wb:
        if flt and not any(f in ws.title.lower() for f in flt):
            continue                              # тек қажетті ай парақтары (мыс. «қазан»)
        cells_rows = list(ws.iter_rows())
        rows = [[c.value for c in cr] for cr in cells_rows]
        h = _find_header(rows, paid_column)
        if not h:
            continue
        hi, ni, pis, qi, li, heads = h
        up.sheets.append(ws.title)
        mcols = [i for i, x in enumerate(heads) if x.lower().startswith("неше айға")]
        acol = next((i for i, x in enumerate(heads) if "төлем суммасы" in x.lower()), None)
        pcol = next((i for i, x in enumerate(heads) if x.lower() == "сумма"), None)
        dcol = next((i for i, x in enumerate(heads) if any(k in x.lower() for k in DATE_KEYS)), None)
        if first:
            up.paid_column, up.level_column, first = heads[qi], heads[li] if li is not None else "", False
        for k, r in enumerate(rows[hi + 1:], start=hi + 2):
            if not r or all(c in (None, "") for c in r):
                continue
            up.total_rows += 1
            mark, mtext = "", ""
            for off in (qi, qi + 1):
                v = r[off] if off < len(r) else None
                t = str(v).strip().lower() if v is not None and not isinstance(v, (int, float)) else ""
                if t and not mtext:
                    mtext = t
                if t in YES:
                    mark = "yes"; break
                if t in ALMOST_YES and not mark:
                    mark = "almost"
            good, bad, raw = [], [], []
            for pi in pis:
                if pi < len(r) and r[pi] not in (None, ""):
                    g, b = parse_phones(r[pi]); good += g; bad += b; raw.append(str(r[pi]))
            name = str(r[ni]).strip() if ni is not None and ni < len(r) and r[ni] else ""
            lvl = str(r[li]).strip() if li is not None and li < len(r) and r[li] else ""
            pdate = None
            if dcol is not None:
                for off in (dcol, dcol + 1):          # қатар бір баған жылжыған болуы мүмкін
                    pdate = parse_date(r[off]) if off < len(r) else None
                    if pdate:
                        break
            amount = 0.0
            if acol is not None:
                for off in (acol, acol + 1):
                    amount = parse_num(r[off]) if off < len(r) else 0.0
                    if amount > 0:
                        break
            plan = parse_num(r[pcol]) if pcol is not None and pcol < len(r) else 0.0
            green = _green_row(cells_rows[k - 1])
            months = 0
            for mi in mcols:
                m = MONTHS_RE.search(str(r[mi])) if mi < len(r) and r[mi] is not None else None
                if m:
                    months = int(m.group(1)); break
            row = UpRow(row=k, sheet=ws.title, name=name, phones=list(dict.fromkeys(good)), bad_phones=bad,
                        raw_phone=" / ".join(raw), level=ws.title, green=green, months=months, marker=mtext,
                        pay_date=pdate, amount=amount, plan=plan)
            if mark == "yes" or green:
                up.paid_rows.append(row)
            elif mark == "almost":
                up.doubtful.append(row)
    if not up.sheets and flt:
        raise UpFileError(f"«{sheet_filter}» деп аталатын парақ табылмады. Файлдағы парақтар: "
                          + ", ".join(w.title for w in wb))
    if not up.sheets:
        raise UpFileError("UP кестесінен аты/телефон және «Келесі айға қалады ма?» бағанын таба алмадым.")
    return up


# ----------------------------------- есеп -----------------------------------
ISSUE_ORDER = ["Жасыл, бірақ белгі қарама-қайшы", "Табылмады", "Нөмір сәйкес емес", "Деңгей сәйкес емес", "Телефон форматы қате",
               "Телефон сәйкес, аты басқа", "Аты бірдей бірнеше оқушы", "UP-та қайталанған", "Белгі түсініксіз"]


@dataclass
class Report:
    by_group: dict
    paid_students: list
    unpaid_students: list
    unmatched_paid: int
    paid_column: str
    ambiguous: int = 0
    issues: list = field(default_factory=list)
    up_paid: int = 0
    up_green: int = 0
    green_students: list = field(default_factory=list)
    payments: list = field(default_factory=list)   # әр төлеген оқушы: күні, куратор, топ, түрі


def build_report(students: list[Student], up: UpData) -> Report:
    by_phone, by_name = defaultdict(list), defaultdict(list)
    for i, s in enumerate(students):
        if s.key_phone:
            by_phone[s.key_phone].append(i)
        by_name[s.key_name].append(i)
    hits, times, green_hits, gmonths, pdates, amts = set(), Counter(), set(), {}, {}, {}
    issues, unmatched, ambiguous = [], 0, 0

    def add(kind, r, st=None, note=""):
        issues.append({
            "Түрі": kind, "UP қатар": f"{r.sheet}:{r.row}", "UP аты": r.name, "UP телефон": r.raw_phone, "UP деңгей": r.level,
            "Платформадағы аты": st.name if st else "", "Платформа нөмірі": st.phone if st else "",
            "Платформа деңгейі": st.level if st else "",
            "Куратор / топ": f"{st.curator} · {st.group}" if st else "", "Түсініктеме": note})

    for r in up.paid_rows:
        m, how = None, ""
        for ph in r.phones:
            if ph in by_phone:
                m, how = by_phone[ph][0], "phone"
                break
        if m is None:
            cand = list(by_name.get(fold_name(r.name), [])) if r.name else []
            lv = canon_level(r.level)
            if lv and len(cand) > 1:
                cand = [i for i in cand if canon_level(students[i].level) == lv] or cand
            if len(cand) == 1:
                m, how = cand[0], "name"
            elif len(cand) > 1:
                ambiguous += 1
                add("Аты бірдей бірнеше оқушы", r, students[cand[0]],
                    "Платформада: " + "; ".join(f"{students[i].phone} ({students[i].curator})" for i in cand)
                    + " — қайсысы екені анықталмады, есепке қосылмады")
                continue
        if m is None:
            unmatched += 1
            best, bi = 0.0, None
            f = fold_name(r.name)
            if f:
                for i, s in enumerate(students):
                    sc = difflib.SequenceMatcher(None, f, s.key_name).ratio()
                    if sc > best:
                        best, bi = sc, i
            note = "Платформа тізімінде жоқ"
            st = None
            if bi is not None and best >= 0.8:
                st = students[bi]
                note = f"Ұқсас оқушы бар ({best:.0%}): {st.name} {st.phone} — есепке қосылмады, тексеріңіз"
            add("Табылмады", r, st, note)
            continue
        s = students[m]
        hits.add(m); times[m] += 1
        if m not in amts or (amts[m][0] == 0 and r.amount > 0):
            amts[m] = (r.amount, r.plan)
        if r.pay_date and (m not in pdates or r.pay_date < pdates[m]):
            pdates[m] = r.pay_date
        if r.green:
            green_hits.add(m); gmonths[m] = max(gmonths.get(m, 0), r.months)
            if r.marker in NEGATIVE:
                add("Жасыл, бірақ белгі қарама-қайшы", r, s,
                    f"Қатар жасыл (бірнеше айға төленген), бірақ белгі «{r.marker}». Төленген деп есептелді — тексеріңіз")
        if how == "name":
            if r.phones:
                add("Нөмір сәйкес емес", r, s, f"UP-та {', '.join(fmt_phone(p) for p in r.phones)}, "
                    f"платформада {s.phone}. Аты бойынша табылды, есепке қосылды")
            elif r.bad_phones:
                add("Телефон форматы қате", r, s, f"«{'; '.join(r.bad_phones)}» нөмір емес. "
                    "Аты бойынша табылды, есепке қосылды")
            else:
                add("Телефон форматы қате", r, s, "UP-та телефон жоқ. Аты бойынша табылды, есепке қосылды")
        else:
            if r.name and difflib.SequenceMatcher(None, fold_name(r.name), s.key_name).ratio() < 0.6:
                add("Телефон сәйкес, аты басқа", r, s, "Нөмір бірдей, бірақ аты мүлде басқа — бөлек адам болуы мүмкін")
        if r.level and canon_level(r.level) and canon_level(s.level) and canon_level(r.level) != canon_level(s.level):
            add("Деңгей сәйкес емес", r, s, f"UP парағы: {r.level} / платформа: {s.level}. Платформадағы топта есептелді")
    for r in up.doubtful:
        add("Белгі түсініксіз", r, None, "«иә» емес, бірақ соған ұқсас мән жазылған — есепке қосылмады. «иә» деп түзетіңіз")
    for i, n in times.items():
        if n > 1:
            s = students[i]
            issues.append({"Түрі": "UP-та қайталанған", "UP қатар": "", "UP аты": s.name, "UP телефон": "",
                           "UP деңгей": "", "Платформадағы аты": s.name, "Платформа нөмірі": s.phone,
                           "Платформа деңгейі": s.level, "Куратор / топ": f"{s.curator} · {s.group}",
                           "Түсініктеме": f"UP-та {n} рет белгіленген, бір рет есептелді"})
    by = defaultdict(lambda: [0, 0, 0]); paid_l, unpaid_l, green_l = [], [], []
    for i, s in enumerate(students):
        k = (s.level, s.curator, s.group)
        by[k][0] += 1
        row = (s.level, s.curator, s.group, s.name, s.phone)
        if i in hits:
            by[k][1] += 1
            paid_l.append(row + ("жасыл (көп ай)" if i in green_hits else "иә",))
            if i in green_hits:
                by[k][2] += 1; green_l.append(row + (gmonths.get(i) or "",))
        else:
            unpaid_l.append(row)
    payments = [{"Күні": pdates.get(i), "Деңгей": students[i].level, "Куратор": students[i].curator,
                 "Топ": students[i].group, "Студент": students[i].name, "Телефон": students[i].phone,
                 "Түрі": "жасыл (көп ай)" if i in green_hits else "иә",
                 "Төленген сумма": amts.get(i, (0, 0))[0], "Баға (СУММА)": amts.get(i, (0, 0))[1]} for i in sorted(hits)]
    issues.sort(key=lambda x: (ISSUE_ORDER.index(x["Түрі"]), str(x["UP қатар"])))
    return Report(dict(by), paid_l, unpaid_l, unmatched, up.paid_column, ambiguous, issues, up.paid_count,
                  up.green_count, green_l, payments)


def _period_counts(rep: Report, start, end, key):
    """Таңдалған күндер аралығында (start..end) төлеген оқушылар саны: key(төлем) бойынша."""
    c = Counter()
    for p in rep.payments:
        d = p["Күні"]
        if d and start <= d <= end:
            c[key(p)] += 1
    return c


def rank_curators(rep: Report, start=None, end=None) -> list[dict]:
    """Кураторлар рейтингі: төлем % (тең болса — төлеген саны көп жоғары).
    start/end берілсе — тек сол күндерде төлегендер (күні көрсетілмегендер есепке кірмейді)."""
    agg = defaultdict(lambda: [0, 0, 0, set(), 0])
    for (lv, cur, grp), (t, p, g) in rep.by_group.items():
        a = agg[cur]; a[0] += t; a[1] += p; a[2] += g; a[3].add(lv); a[4] += 1
    per = _period_counts(rep, start, end, lambda p: p["Куратор"]) if start else None
    out = []
    for c, a in agg.items():
        paid = per.get(c, 0) if per is not None else a[1]
        out.append({"Куратор": c, "Деңгей": ", ".join(sorted(a[3])), "Топ саны": a[4], "Оқушы": a[0], "Төледі": paid,
                    "🟩": a[2] if per is None else sum(1 for p in rep.payments if p["Куратор"] == c
                                                       and p["Күні"] and start <= p["Күні"] <= end
                                                       and p["Түрі"] != "иә"),
                    "%": round(paid / a[0] * 100, 1) if a[0] else 0.0})
    out.sort(key=lambda x: (-x["%"], -x["Төледі"], x["Куратор"]))
    for i, x in enumerate(out, 1):
        x["№"] = i
    return out


def rank_groups(rep: Report, start=None, end=None) -> list[dict]:
    per = _period_counts(rep, start, end, lambda p: (p["Деңгей"], p["Куратор"], p["Топ"])) if start else None
    out = []
    for (lv, c, g), (t, p, gg) in rep.by_group.items():
        paid = per.get((lv, c, g), 0) if per is not None else p
        out.append({"Топ": short_group(c, g), "Деңгей": lv, "Оқушы": t, "Төледі": paid,
                    "🟩": gg if per is None else sum(1 for x in rep.payments if (x["Деңгей"], x["Куратор"], x["Топ"]) == (lv, c, g)
                                                      and x["Күні"] and start <= x["Күні"] <= end and x["Түрі"] != "иә"),
                    "%": round(paid / t * 100, 1) if t else 0.0})
    out.sort(key=lambda x: (-x["%"], -x["Төледі"], x["Топ"]))
    for i, x in enumerate(out, 1):
        x["№"] = i
    return out


def daily_counts(rep: Report) -> dict:
    """Күн бойынша төлеген оқушы саны + күні көрсетілмегендер (None кілтімен)."""
    c = Counter(p["Күні"] for p in rep.payments)
    return dict(sorted(c.items(), key=lambda kv: (kv[0] is None, kv[0] or dt.date.min)))


def level_analytics(rep: Report) -> list[dict]:
    """Деңгей бойынша retention: оқушы / төледі / бірнеше айға (жасыл) / тек бір айға («иә») / төлемеген."""
    tot, paid, green = Counter(), Counter(), Counter()
    for (lv, _c, _g), (t, p, g) in rep.by_group.items():
        tot[lv] += t; paid[lv] += p; green[lv] += g
    m2, m3, mx = Counter(), Counter(), Counter()
    for row in rep.green_students:                    # (деңгей, куратор, топ, аты, тел, ай саны)
        n = row[5] if isinstance(row[5], int) else 0
        (m2 if n == 2 else m3 if n >= 3 else mx)[row[0]] += 1
    dated, undated = Counter(), Counter()
    for p in rep.payments:
        (dated if p["Күні"] else undated)[p["Деңгей"]] += 1
    out = []
    for lv in sorted(tot):
        t, p, g = tot[lv], paid[lv], green[lv]
        out.append({"Деңгей": lv, "Оқушы": t, "Төледі": p, "Retention %": round(p / t * 100, 1) if t else 0.0,
                    "Бірнеше айға (🟩)": g, "оның ішінде 2 ай": m2[lv], "3+ ай": m3[lv], "ай саны жоқ": mx[lv],
                    "Тек 1 айға («иә»)": p - g, "Төлемеген": t - p,
                    "Күні бар төлем": dated[lv], "Күні жоқ төлем": undated[lv]})
    return out


def daily_by_level(rep: Report) -> list[dict]:
    """Деңгей бойынша күн сайынғы жаңа төлем және жиынтық retention (тек күні көрсетілген төлемдер)."""
    dates = [p["Күні"] for p in rep.payments if p["Күні"]]
    if not dates:
        return []
    tot = Counter()
    for (lv, _c, _g), (t, _p, _gg) in rep.by_group.items():
        tot[lv] += t
    new = Counter((p["Деңгей"], p["Күні"]) for p in rep.payments if p["Күні"])
    # График қыркүйектен басталады (оқу маусымы). Маусым жылы бүгінгі күннен анықталады: қыркүйектен кейін — осы жыл,
    # одан бұрын — өткен жыл. Ай өткен сайын график өзі ұзарады, ескі жылдың күндері көрінбейді.
    today = dt.date.today()
    d0 = dt.date(today.year if today.month >= 9 else today.year - 1, 9, 1)
    d1 = max(max(dates), today)
    days = [d0 + dt.timedelta(days=i) for i in range((d1 - d0).days + 1)]
    out = []
    for lv in sorted(tot):
        cum = sum(n for (l, d), n in new.items() if l == lv and d < d0)      # қыркүйектен бұрынғы төлемдер бастапқы сан
        for d in days:
            cum += new.get((lv, d), 0)
            out.append({"Күні": d, "Деңгей": lv, "Жаңа төлем": new.get((lv, d), 0), "Жиынтық": cum,
                        "Retention %": round(cum / tot[lv] * 100, 1) if tot[lv] else 0.0})
    return out


def short_group(cur: str, group: str) -> str:
    m = re.search(r"(\d+)-ТОБЫ", group or "")
    return f"{cur} · {m.group(1)}-топ" if m else cur


def format_text(rep: Report) -> str:
    lines, tot, paid, gr = [], 0, 0, 0
    for lv in sorted({k[0] for k in rep.by_group}):
        items = [(k, v) for k, v in rep.by_group.items() if k[0] == lv]
        lt, lp, lg = (sum(v[j] for _, v in items) for j in range(3))
        lines.append(f"\n📚 {lv}: {lp}/{lt} ({lp / lt * 100:.0f}%)" + (f" · 🟩 {lg}" if lg else ""))
        for (l, cur, grp), (t, p, g) in sorted(items, key=lambda x: (-x[1][1] / max(x[1][0], 1), x[0][1], x[0][2])):
            lines.append(f"• {short_group(cur, grp)}: {p}/{t} ({p / t * 100:.0f}%)" + (f" · 🟩 {g}" if g else ""))
        tot += lt; paid += lp; gr += lg
    head = (f"📊 Келесі айға төлем есебі\nЖАЛПЫ: {paid}/{tot} ({paid / tot * 100:.1f}%)\n"
            f"оның ішінде 🟩 бірнеше айға төлегендер (жасыл қатар): {gr}\n"
            f"«иә» деп белгіленгендер (1 ай): {paid - gr}\n"
            f"UP-тағы төленген қатар саны: {rep.up_paid}") if tot else "Тізім бос"
    rk = rank_curators(rep)
    medals = ["🥇", "🥈", "🥉", "4.", "5."]
    top = "\n\n🏆 Топ-5 кураторлар (барлық топтары бойынша)\n" + "\n".join(
        f"{medals[i]} {x['Куратор']}: {x['Төледі']}/{x['Оқушы']} ({x['%']:.0f}%)" for i, x in enumerate(rk[:5]))
    foot = (top if rk else "") + f"\n\n(төлем бағаны: «{rep.paid_column}»)"
    if rep.issues:
        c = Counter(i["Түрі"] for i in rep.issues)
        foot += f"\n\n⚠️ Мәселелер: {len(rep.issues)} (толық тізім Excel-де, «Мәселелер» парағы)"
        for k in ISSUE_ORDER:
            if c.get(k):
                foot += f"\n• {k}: {c[k]}"
    return head + "\n" + "\n".join(lines) + foot


def write_report_xlsx(rep: Report, path) -> Path:
    wb = openpyxl.Workbook()
    H, HF = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78")

    def sh(title, head, data, widths=None, first=False):
        s = wb.active if first else wb.create_sheet()
        s.title = title; s.append(head)
        for c in s[1]:
            c.font, c.fill, c.alignment = H, HF, Alignment(wrap_text=True)
        for d in data:
            s.append(list(d))
        s.freeze_panes = "A2"
        for i in range(1, len(head) + 1):
            s.column_dimensions[openpyxl.utils.get_column_letter(i)].width = (widths or {}).get(i, 24)
        return s
    rows = [[l, c, g, t, p, f"=IF(D{i}=0,0,E{i}/D{i})", gg, p - gg] for i, ((l, c, g), (t, p, gg)) in
            enumerate(sorted(rep.by_group.items()), start=2)]
    s = sh("Қорытынды", ["Деңгей", "Куратор", "Топ", "Оқушы саны", "Төледі (барлығы)", "%",
                         "оның ішінде жасыл (көп ай)", "«иә» (1 ай)"], rows,
           {1: 22, 2: 24, 3: 42}, first=True)
    for i in range(2, len(rows) + 2):
        s.cell(i, 6).number_format = "0%"
    n = len(rows) + 2
    s.cell(n, 1, "ЖАЛПЫ").font = Font(bold=True)
    s.cell(n, 4, f"=SUM(D2:D{n-1})"); s.cell(n, 5, f"=SUM(E2:E{n-1})"); s.cell(n, 6, f"=IF(D{n}=0,0,E{n}/D{n})")
    s.cell(n, 6).number_format = "0%"
    s.cell(n, 7, f"=SUM(G2:G{n-1})"); s.cell(n, 8, f"=SUM(H2:H{n-1})")
    cols = ["Түрі", "UP қатар", "UP аты", "UP телефон", "UP деңгей", "Платформадағы аты", "Платформа нөмірі",
            "Платформа деңгейі", "Куратор / топ", "Түсініктеме"]
    sh("Мәселелер", cols, [[i[c] for c in cols] for i in rep.issues],
       {1: 26, 3: 24, 4: 20, 6: 24, 7: 16, 8: 20, 9: 36, 10: 70})
    sh("Төлегендер", ["Деңгей", "Куратор", "Топ", "Студент", "Телефон", "Түрі"], rep.paid_students, {3: 42})
    la = level_analytics(rep)
    if la:
        sh("Retention (деңгей)", list(la[0].keys()), [list(x.values()) for x in la], {1: 24})
    dl = daily_by_level(rep)
    if dl:
        sh("Retention (күндер)", ["Күні", "Деңгей", "Жаңа төлем", "Жиынтық", "Retention %"],
           [[x["Күні"], x["Деңгей"], x["Жаңа төлем"], x["Жиынтық"], x["Retention %"]] for x in dl], {1: 14, 2: 24})
    sh("Жасыл (көп ай)", ["Деңгей", "Куратор", "Топ", "Студент", "Телефон", "Неше ай"], rep.green_students, {3: 42})
    sh("Төлемегендер", ["Деңгей", "Куратор", "Топ", "Студент", "Телефон"], rep.unpaid_students, {3: 42})
    wb.save(path)
    return Path(path)
