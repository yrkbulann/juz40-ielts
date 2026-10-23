"""JUZ40 IELTS СЫРТҚЫ: платформадағы төлемдер файлынан («Төлемдер» парағы) келесі айға (Қараша) төлегендер."""
import re
import datetime as dt
from dataclasses import dataclass
from collections import defaultdict
from pathlib import Path
import openpyxl
import core

NEXT_MONTH = "қараша"      # келесі ай: курс атауында осы сөз болса — келесі айға төленген
MONTHS_KZ = {"қаңтар": 1, "ақпан": 2, "наурыз": 3, "сәуір": 4, "мамыр": 5, "маусым": 6, "шілде": 7,
             "тамыз": 8, "қыркүйек": 9, "қазан": 10, "қараша": 11, "желтоқсан": 12}
DATE_KZ = re.compile(r"(\d{4})\s*ж\.\s*(\d{1,2})\s+([^\W\d_]+)(?:,\s*(\d{1,2}):(\d{2}))?")


@dataclass
class ExtPayment:
    name: str
    phone: str
    email: str
    city: str
    course: str          # толық атауы
    track: str           # General English / IELTS
    product: str         # ELEM 4.0 | ҚАРАША
    level: str           # ELEM / PRE / BEGIN / IELTS / БАСҚА
    months: int
    amount: float
    method: str
    date: object         # datetime.date | None
    manager: str
    status: str
    prepay: bool
    known: str = ""      # куратор базасында бар болса — «куратор · топ»


def parse_kz_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    m = DATE_KZ.search(str(v or ""))
    if not m:
        return core.parse_date(v)
    mo = MONTHS_KZ.get(m.group(3).lower())
    if not mo:
        return None
    try:
        return dt.date(int(m.group(1)), mo, int(m.group(2)))
    except ValueError:
        return None


def _months(v) -> int:
    nums = [int(x) for x in re.findall(r"\d+", str(v or ""))]
    return sum(nums) if nums else 0


def load_payments(path, next_month: str = NEXT_MONTH) -> list[ExtPayment]:
    ws = openpyxl.load_workbook(path, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    head = [str(h or "").strip().lower() for h in rows[0]]

    def col(*names):
        for i, h in enumerate(head):
            if any(n == h for n in names):
                return i
        for i, h in enumerate(head):
            if any(n in h for n in names):
                return i
        return None
    c = {k: col(*v) for k, v in {
        "name": ("студент",), "phone": ("телефон",), "email": ("email",), "city": ("қала",),
        "course": ("курстар",), "months": ("төленген ай саны",), "amount": ("сомасы",), "method": ("әдіс",),
        "date": ("күні",), "seller": ("сатушы менеджер",), "filler": ("толтырған",), "status": ("күйі",)}.items()}
    if c["course"] is None or c["name"] is None:
        raise core.UpFileError("Төлемдер файлынан «Студент» және «Курстар» бағандары табылмады.")

    def get(r, k):
        i = c[k]
        return r[i] if i is not None and i < len(r) else None
    out = []
    for r in rows[1:]:
        course = str(get(r, "course") or "").strip()
        if not course or next_month.lower() not in course.lower():
            continue
        if course.lower().startswith("qa"):
            continue                                   # сынақ жазбалар
        status = str(get(r, "status") or "").strip()
        track, _, rest = course.partition("·")
        product = rest.split("—", 1)[1].strip() if "—" in rest else rest.strip()
        g, _b = core.parse_phones(get(r, "phone"))
        mgr = str(get(r, "seller") or "").strip()
        if mgr in ("", "—", "Менеджерсіз"):
            mgr = str(get(r, "filler") or "").strip()
        if mgr in ("", "—"):
            mgr = "Менеджерсіз"
        lvl = core.canon_level(product) or core.canon_level(course) or "БАСҚА"
        out.append(ExtPayment(
            name=str(get(r, "name") or "").strip(), phone=core.fmt_phone(g[0]) if g else str(get(r, "phone") or ""),
            email=str(get(r, "email") or "").replace("—", "").strip(), city=str(get(r, "city") or "").replace("—", "").strip(),
            course=course, track=track.strip(), product=product, level=lvl, months=_months(get(r, "months")),
            amount=core.parse_num(re.sub(r"[^\d.,]", "", str(get(r, "amount") or ""))), method=str(get(r, "method") or "").strip(),
            date=parse_kz_date(get(r, "date")), manager=mgr, status=status,
            prepay="предоплата" in course.lower()))
    return [p for p in out if p.status in ("", "Төленді")]


def mark_known(pays: list[ExtPayment], students: list) -> None:
    """Куратор базасында бар оқушыны белгілейді (телефон, болмаса аты бойынша)."""
    byp, byn = {}, defaultdict(list)
    for s in students:
        if s.key_phone:
            byp.setdefault(s.key_phone, s)
        byn[s.key_name].append(s)
    for p in pays:
        g, _ = core.parse_phones(p.phone)
        s = byp.get(g[0]) if g else None
        if s is None:
            cand = byn.get(core.fold_name(p.name), [])
            s = cand[0] if len(cand) == 1 else None
        if s is not None:
            p.known = f"{s.curator} · {core.short_group(s.curator, s.group).split(' · ')[-1]}"


def students_table(pays) -> list[dict]:
    return [{"Күні": p.date, "Оқушы": p.name, "Телефон": p.phone, "Курс": p.product, "Деңгей": p.level,
             "Ай саны": p.months, "Сома": p.amount, "Төлем түрі": "алдын ала (предоплата)" if p.prepay else "толық",
             "Менеджер": p.manager, "Қала": p.city, "Базада бар (куратор)": p.known} for p in
            sorted(pays, key=lambda x: (x.date or dt.date.min, x.name), reverse=True)]


def summary(pays) -> dict:
    uniq = {p.phone or p.name for p in pays}
    return {"Төлем саны": len(pays), "Оқушы (бірегей)": len(uniq), "Жалпы сома": sum(p.amount for p in pays),
            "1 айға": sum(1 for p in pays if p.months == 1), "2+ айға": sum(1 for p in pays if p.months >= 2),
            "Базада бар": sum(1 for p in pays if p.known)}


def by_key(pays, keyfn, label) -> list[dict]:
    d = defaultdict(lambda: [0, 0.0, 0, 0])
    for p in pays:
        a = d[keyfn(p)]; a[0] += 1; a[1] += p.amount; a[2] += p.months == 1; a[3] += p.months >= 2
    out = [{label: k, "Оқушы": v[0], "Сома": v[1], "1 айға": v[2], "2+ айға": v[3]} for k, v in d.items()]
    out.sort(key=lambda x: (-x["Оқушы"], -x["Сома"], str(x[label])))
    for i, x in enumerate(out, 1):
        x["№"] = i
    return out


def daily(pays) -> list[dict]:
    dated = [p for p in pays if p.date]
    if not dated:
        return []
    d0, d1 = min(p.date for p in dated), max(p.date for p in dated)
    cnt, amt = defaultdict(int), defaultdict(float)
    for p in dated:
        cnt[p.date] += 1; amt[p.date] += p.amount
    out, cum, cs = [], 0, 0.0
    for i in range((d1 - d0).days + 1):
        d = d0 + dt.timedelta(days=i)
        cum += cnt[d]; cs += amt[d]
        out.append({"Күні": d, "Жаңа төлем": cnt[d], "Жиынтық": cum, "Сома": amt[d], "Жиынтық сома": cs})
    return out
