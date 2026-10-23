"""Админ отчеттері: JUZ40 IELTS форматындағы күндік және жалпы мәтіндік отчеттер."""
import datetime as dt
from collections import Counter
import core

# Отчеттағы топтар: (атауы, {(деңгей, ай), …}). Жаңа топ қосу — осы тізімге жол қосу.
ISHKI = [("IELTS ҚЫРКҮЙЕК", {("IELTS", "ҚЫРКҮЙЕК")}),
         ("IELTS ҚАЗАН", {("IELTS", "ҚАЗАН")}),
         ("GE ҚАЗАН", {("ELEM", "ҚАЗАН"), ("PRE", "ҚАЗАН")})]
LEVEL_TO_LEVEL = [("ELEM ҚЫРКҮЙЕК", {("ELEM", "ҚЫРКҮЙЕК")}),
                  ("PRE ҚЫРКҮЙЕК", {("PRE", "ҚЫРКҮЙЕК")}),
                  ("BEGINNER ҚАЗАН", {("BEGIN", "ҚАЗАН")})]
TOTAL_LABEL = {"GE ҚАЗАН": "GE(ELEM/PRE) ҚАЗАН"}       # ЖАЛПЫ отчетте басқаша жазылады
MONTHS = (("ҚЫРКҮЙЕК", "ҚЫРКҮЙЕК"), ("QYRKUIEK", "ҚЫРКҮЙЕК"), ("ҚАЗАН", "ҚАЗАН"), ("QAZAN", "ҚАЗАН"))


def level_key(level: str):
    up = (level or "").upper()
    month = next((m for w, m in MONTHS if w in up), "")
    return core.canon_level(level), month


def _which(groups, level):
    k = level_key(level)
    for name, keys in groups:
        if k in keys:
            return name
    return None


def pct(p, t) -> str:
    return f"{(p / t * 100 if t else 0):.2f}".replace(".", ",") + "%"


def _students(rep, groups):
    c = Counter()
    for (lv, _c, _g), (t, _p, _gg) in rep.by_group.items():
        n = _which(groups, lv)
        if n:
            c[n] += t
    return c


def daily_report(rep, day: dt.date):
    allg = ISHKI + LEVEL_TO_LEVEL
    c = Counter()
    for p in rep.payments:
        if p["Күні"] == day:
            n = _which(allg, p["Деңгей"])
            if n:
                c[n] += 1
    ds = day.strftime("%d.%m")
    L = [f"JUZ40 IELTS | ІШКІ САТУ {ds} КҮНДІК ОТЧЕТ", "",
         f"{ISHKI[0][0]}: {c[ISHKI[0][0]]}", "",
         f"{ISHKI[1][0]}: {c[ISHKI[1][0]]}", f"{ISHKI[2][0]}: {c[ISHKI[2][0]]}", "", "LL: "]
    L += [f"{n}: {c[n]}" for n, _ in LEVEL_TO_LEVEL]
    L += ["", f"ЖАЛПЫ (оқушы): {sum(c[n] for n, _ in allg)}"]
    return "\n".join(L)


def _total_counts(rep, groups, start, end):
    paid = Counter()
    for p in rep.payments:
        d = p["Күні"]
        if d is not None and ((start and d < start) or d > end):
            continue                      # күні бар, бірақ кезеңнен тыс
        n = _which(groups, p["Деңгей"])
        if n:
            paid[n] += 1                  # күні жоқ төлемдер (жасыл) жалпы санға кіреді
    return paid


def total_reports(rep, start: dt.date, end: dt.date):
    """(ІШКІ САТУ жалпы отчеті, LEVEL TO LEVEL жалпы отчеті)."""
    rng = f"{start.strftime('%d.%m')}-{end.strftime('%d.%m')}"
    out = []
    pi, ti = _total_counts(rep, ISHKI, start, end), _students(rep, ISHKI)
    L = [f"JUZ40 IELTS | ІШКІ САТУ {rng} ЖАЛПЫ ОТЧЕТ", ""]
    for i, (n, _) in enumerate(ISHKI):
        if i == 1:
            L.append("")
        L.append(f"{TOTAL_LABEL.get(n, n)}: {pi[n]}/{ti[n]}({pct(pi[n], ti[n])})")
    P, T = sum(pi.values()), sum(ti.values())
    L += ["", f"ЖАЛПЫ: {P}/{T} ({pct(P, T)})"]
    out.append("\n".join(L))
    pl, tl = _total_counts(rep, LEVEL_TO_LEVEL, start, end), _students(rep, LEVEL_TO_LEVEL)
    L = [f"JUZ40 IELTS | LEVEL TO LEVEL {rng} ЖАЛПЫ ОТЧЕТ", ""]
    L += [f"{n}: {pl[n]}/{tl[n]}({pct(pl[n], tl[n])})" for n, _ in LEVEL_TO_LEVEL]
    out.append("\n".join(L))
    return out[0], out[1]


def report_warnings(rep) -> list[str]:
    """Отчет дұрыс шығуына кедергі жағдайлар (мысалы, базада жоқ топтар)."""
    w = []
    t = _students(rep, ISHKI + LEVEL_TO_LEVEL)
    for n, _ in ISHKI + LEVEL_TO_LEVEL:
        if t[n] == 0:
            w.append(f"«{n}» оқушылары базада жоқ — отчетте 0 болып шығады.")
    und = sum(1 for p in rep.payments if p["Күні"] is None)
    if und:
        w.append(f"{und} төлемнің күні UP-та жоқ (көбі жасыл қатарлар): олар жалпы отчетке кіреді, күндік отчетке кірмейді.")
    return w


# ------------------------------------------------------------------ СЫРТҚЫ САТУ отчеттері
LEVEL_NAME = {"BEGIN": "BEGINNER", "ELEM": "ELEMENTARY", "PRE": "PRE-INTER"}


def _pct_short(p, t) -> str:
    """31% (бүтін болса), әйтпесе 17,27%."""
    v = p / t * 100 if t else 0
    return (f"{v:.0f}" if abs(v - round(v)) < 0.005 else f"{v:.2f}".replace(".", ",")) + "%"


def external_daily_report(pays, day: dt.date) -> str:
    """JUZ40 IELTS | СЫРТҚЫ САТУ күндік отчеті. pays — толық төлемдер (предоплатасыз)."""
    day_p = [p for p in pays if p.date == day]
    ielts = sum(1 for p in day_p if p.level == "IELTS")
    ge = sum(1 for p in day_p if p.level in LEVEL_NAME)
    return "\n".join([
        f"JUZ40 IELTS | СЫРТҚЫ САТУ {day.strftime('%d.%m')} КҮНДІК ОТЧЕТ", "",
        f"IELTS (оқушы): {ielts}", f"GE (оқушы): {ge}", "",
        f"ЖАЛПЫ (оқушы): {ielts + ge}", "",
        "КОНВЕРСИЯ ЖАҢА ЛИДТАН: 0%", "КОНВЕРСИЯ ЖАҢА БАЗАДАН: 0%"])


def external_total_report(pays, start: dt.date, end: dt.date, plan_ielts: int = 0, plan_ge: int = 0,
                          month_label: str = "ҚАРАША") -> str:
    """СЫРТҚЫ САТУ жалпы отчеті. plan_* — «/»-дан кейінгі сан (қолмен енгізіледі)."""
    sel = [p for p in pays if p.date and start <= p.date <= end]
    ielts = sum(1 for p in sel if p.level == "IELTS")
    by_lv = {k: sum(1 for p in sel if p.level == k) for k in LEVEL_NAME}
    ge = sum(by_lv.values())

    def frac(n, plan):
        return f"{n}/{plan} ({_pct_short(n, plan)})" if plan else f"{n}"
    L = [f"JUZ40 IELTS | СЫРТҚЫ САТУ {start.strftime('%d.%m')}-{end.strftime('%d.%m')} ЖАЛПЫ ОТЧЕТ", ""]
    L += [f"{LEVEL_NAME[k]}: {by_lv[k]}" for k in ("BEGIN", "ELEM", "PRE")]
    L += ["", f"IELTS {month_label} (оқушы): {frac(ielts, plan_ielts)}",
          f"GE {month_label} (оқушы): {frac(ge, plan_ge)}", "",
          f"ЖАЛПЫ (оқушы): {frac(ielts + ge, plan_ielts + plan_ge if plan_ielts and plan_ge else 0)}"]
    return "\n".join(L)
