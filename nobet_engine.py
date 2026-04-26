import calendar
import random
from datetime import date, timedelta, datetime
from dataclasses import dataclass
from collections import defaultdict

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font


# =========================================================
# AYARLAR
# =========================================================

@dataclass
class Settings:
    MAX_SAME_WEEKDAY: int = 2
    WEEKDAY_PENALTY: float = 1.5
    DENGE_KATSAYI: float = 0.7
    AGIR_GUN_FRENI: float = 0.4
    MIN_GAP_DAYS: int = 14

    MAX_GAP_DAYS: int = 35
    MAX_GAP_BONUS: float = 6
    MAX_GAP_EXTRA_PER_DAY: float = 0.15

    IDEAL_WEEKEND_RATIO: float = 2 / 7
    WEEKEND_BALANCE_PENALTY: float = 42
    WEEKDAY_BALANCE_PENALTY: float = 28
    EXTREME_WEEKEND_PENALTY: float = 55
    BALANCE_POWER: float = 2.6
    MIN_TOTAL_FOR_BALANCE: int = 3

    COUNT_GAP_PENALTY_WEEKEND: float = 16
    COUNT_GAP_PENALTY_WEEKDAY: float = 10

    MONTHLY_ZERO_ASSIGNMENT_BONUS: float = 2.40
    MONTHLY_TWO_PLUS_PENALTY: float = 1.10
    LATE_MONTH_EXTRA_ZERO_BONUS: float = 2.80
    VERY_LATE_MONTH_EXTRA_ZERO_BONUS: float = 1.60
    LATE_MONTH_RATIO: float = 0.70
    VERY_LATE_MONTH_RATIO: float = 0.85

    RECENT_BAYRAM_WINDOW_DAYS: int = 365
    VERY_RECENT_BAYRAM_WINDOW_DAYS: int = 180
    RECENT_BAYRAM_PENALTY: float = 18
    VERY_RECENT_BAYRAM_PENALTY: float = 26
    SAME_BAYRAM_TYPE_RECENT_PENALTY: float = 12
    AREFE_RECENT_PENALTY: float = 8


SETTINGS = Settings()


# =========================================================
# GLOBAL STATE
# =========================================================

monthly_stats = defaultdict(lambda: defaultdict(lambda: {
    "bayram": 0,
    "arefe": 0,
    "pazartesi": 0,
    "salı": 0,
    "çarşamba": 0,
    "perşembe": 0,
    "cuma": 0,
    "cumartesi": 0,
    "pazar": 0,
}))

eklenme_tarihi = {}
cikma_tarihi = {}
eklenme = {}

# İstersen buraya bayram tutmasını istemediğin eczaneleri yazabilirsin.
BAYRAM_YASAK = {
    "ECZANE1",
    "ECZANE5",
}


# =========================================================
# İSİM NORMALİZASYONU
# =========================================================

NAME_ALIASES = {}


def normalize_name(name):
    if name is None:
        return ""
    name = str(name).strip().upper()
    return NAME_ALIASES.get(name, name)


def safe_int(v):
    try:
        if v is None:
            return 0
        if pd.isna(v):
            return 0
        if str(v).strip() == "":
            return 0
        return int(float(v))
    except Exception:
        return 0


# =========================================================
# EXCEL / DATAFRAME OKUYUCU
# =========================================================

DAY_KEYS = {
    0: "pazartesi",
    1: "salı",
    2: "çarşamba",
    3: "perşembe",
    4: "cuma",
    5: "cumartesi",
    6: "pazar",
}

DAY_SHORT = ["Pzt", "Salı", "Çarş", "Perş", "Cuma", "Ctesi", "Pazar"]


def excelden_gecmis_yuk_df(df):
    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]

    kolon_map = {
        "eczane": None,
        "bayram": None,
        "pzt": None,
        "salı": None,
        "çarş": None,
        "perş": None,
        "cuma": None,
        "ctesi": None,
        "pazar": None,
    }

    for c in df.columns:
        if "eczane" in c:
            kolon_map["eczane"] = c
        elif "bayram" in c:
            kolon_map["bayram"] = c
        elif "pzt" in c or "pazartesi" in c:
            kolon_map["pzt"] = c
        elif "sal" in c:
            kolon_map["salı"] = c
        elif "çar" in c:
            kolon_map["çarş"] = c
        elif "per" in c:
            kolon_map["perş"] = c
        elif "cuma" in c and "cumartesi" not in c:
            kolon_map["cuma"] = c
        elif "cte" in c or "cumartesi" in c:
            kolon_map["ctesi"] = c
        elif "pazar" in c:
            kolon_map["pazar"] = c

    if kolon_map["eczane"] is None:
        raise Exception("Geçmiş yük dosyasında ECZANE kolonu bulunamadı")

    gecmis_yuk = {}

    for _, r in df.iterrows():
        eczane_raw = r.get(kolon_map["eczane"])
        if eczane_raw is None or pd.isna(eczane_raw):
            continue

        eczane = normalize_name(eczane_raw)
        if not eczane:
            continue

        gecmis_yuk[eczane] = {
            "bayram": safe_int(r.get(kolon_map["bayram"], 0)),
            "pazartesi": safe_int(r.get(kolon_map["pzt"], 0)),
            "salı": safe_int(r.get(kolon_map["salı"], 0)),
            "çarşamba": safe_int(r.get(kolon_map["çarş"], 0)),
            "perşembe": safe_int(r.get(kolon_map["perş"], 0)),
            "cuma": safe_int(r.get(kolon_map["cuma"], 0)),
            "cumartesi": safe_int(r.get(kolon_map["ctesi"], 0)),
            "pazar": safe_int(r.get(kolon_map["pazar"], 0)),
        }

    return gecmis_yuk


def excelden_gecmis_yuk(excel_dosya):
    return excelden_gecmis_yuk_df(pd.read_excel(excel_dosya))


def parse_excel_date(v):
    if v is None:
        return None
    if isinstance(v, pd.Timestamp):
        return v.date()
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v

    text = str(v).strip()
    if not text:
        return None

    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y.%m.%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except Exception:
            pass

    try:
        return pd.to_datetime(text, dayfirst=True).date()
    except Exception:
        return None


def normalize_bayram_type(text):
    t = str(text).strip().lower()
    if "arefe" in t:
        return "AREFE"
    if "dini" in t:
        return "DINI"
    if "resmi" in t:
        return "RESMI"
    return "BAYRAM"


def excelden_gecmis_bayram_df(df):
    if df is None or len(df) == 0:
        return []

    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]

    col_tarih = None
    col_tur = None
    col_eczane = None

    for c in df.columns:
        if "tarih" in c:
            col_tarih = c
        elif "tur" in c or "tür" in c:
            col_tur = c
        elif "eczane" in c:
            col_eczane = c

    if col_tarih is None or col_eczane is None:
        return []

    kayitlar = []

    for _, r in df.iterrows():
        d = parse_excel_date(r.get(col_tarih))
        e = r.get(col_eczane)

        if d is None or e is None or pd.isna(e):
            continue

        eczane = normalize_name(e)
        if not eczane:
            continue

        tur = normalize_bayram_type(r.get(col_tur, "BAYRAM"))

        kayitlar.append({
            "tarih": d,
            "tur": tur,
            "eczane": eczane,
        })

    return kayitlar


# =========================================================
# TATİL / AREFE
# =========================================================

def turkiye_tatilleri(year):
    return {
        date(year, 1, 1),
        date(year, 4, 23),
        date(year, 5, 1),
        date(year, 5, 19),
        date(year, 7, 15),
        date(year, 8, 30),
        date(year, 10, 29),

        date(2026, 3, 20), date(2026, 3, 21), date(2026, 3, 22),
        date(2026, 5, 27), date(2026, 5, 28), date(2026, 5, 29), date(2026, 5, 30),

        date(2027, 3, 20), date(2027, 3, 21), date(2027, 3, 22),
        date(2027, 5, 27), date(2027, 5, 28), date(2027, 5, 29), date(2027, 5, 30),
    }


def arefe_gunleri(year):
    return {
        date(2026, 3, 19),
        date(2026, 5, 26),
        date(2027, 3, 19),
        date(2027, 5, 26),
    }


def bayram_turu_of_date(d):
    if d in arefe_gunleri(d.year):
        return "AREFE"

    dini = {
        date(2026, 3, 20), date(2026, 3, 21), date(2026, 3, 22),
        date(2026, 5, 27), date(2026, 5, 28), date(2026, 5, 29), date(2026, 5, 30),
        date(2027, 3, 20), date(2027, 3, 21), date(2027, 3, 22),
        date(2027, 5, 27), date(2027, 5, 28), date(2027, 5, 29), date(2027, 5, 30),
    }

    if d in dini:
        return "DINI"

    resmi = {
        date(d.year, 1, 1),
        date(d.year, 4, 23),
        date(d.year, 5, 1),
        date(d.year, 5, 19),
        date(d.year, 7, 15),
        date(d.year, 8, 30),
        date(d.year, 10, 29),
    }

    if d in resmi:
        return "RESMI"

    return "BAYRAM"


# =========================================================
# GÜN KATSAYISI
# Not: İkinci koddaki mantığa göre bayram/arefe katsayısı 0 tutuldu.
# =========================================================

def day_weight(d, tatil, arefe):
    if d in tatil or d in arefe:
        return 0
    if d.weekday() == 6:
        return 2.0
    if d.weekday() == 5:
        return 1.5
    return 1.0


# =========================================================
# AKTİFLİK VE YARDIMCILAR
# =========================================================

def is_person_active_on_date(p, d):
    if p in eklenme_tarihi and d < eklenme_tarihi[p]:
        return False
    if p in cikma_tarihi and d >= cikma_tarihi[p]:
        return False
    return True


def month_progress_ratio(d, month_total_days):
    if month_total_days <= 0:
        return 0
    return d.day / month_total_days


def get_gap_days(last_dates, p, d):
    if p not in last_dates:
        return 999
    return (d - last_dates[p]).days


def get_weekend_count(weekday_stats, p):
    return weekday_stats[p][5] + weekday_stats[p][6]


def get_weekday_count(weekday_stats, p):
    return sum(weekday_stats[p][i] for i in range(5))


def get_last_bayram_gap_days(bayram_dates, p, d):
    dates = bayram_dates.get(p, [])
    valid = [x for x in dates if d >= x]
    if not valid:
        return 9999
    return min((d - x).days for x in valid)


def get_last_arefe_gap_days(arefe_dates, p, d):
    dates = arefe_dates.get(p, [])
    valid = [x for x in dates if d >= x]
    if not valid:
        return 9999
    return min((d - x).days for x in valid)


def count_recent_bayram_same_type(bayram_dates, p, d, target_type, window_days):
    cnt = 0
    for bd in bayram_dates.get(p, []):
        gap = (d - bd).days
        if 0 <= gap <= window_days and bayram_turu_of_date(bd) == target_type:
            cnt += 1
    return cnt


# =========================================================
# SKOR HESABI
# =========================================================

def score_weekend_balance(p, d, weekday_stats):
    hafta_ici = get_weekday_count(weekday_stats, p)
    hafta_sonu = get_weekend_count(weekday_stats, p)
    toplam = hafta_ici + hafta_sonu

    if toplam < SETTINGS.MIN_TOTAL_FOR_BALANCE:
        return 0

    ideal_hafta_sonu = toplam * SETTINGS.IDEAL_WEEKEND_RATIO
    oran = hafta_sonu / toplam if toplam > 0 else 0
    oran_farki = oran - SETTINGS.IDEAL_WEEKEND_RATIO
    adet_farki = hafta_sonu - ideal_hafta_sonu

    skor = 0

    if d.weekday() >= 5:
        if oran_farki > 0:
            skor += SETTINGS.WEEKEND_BALANCE_PENALTY * (abs(oran_farki) ** SETTINGS.BALANCE_POWER) * (toplam + 2)

        if adet_farki > 0:
            skor += SETTINGS.COUNT_GAP_PENALTY_WEEKEND * (adet_farki ** 2)

        if hafta_sonu >= max(2, hafta_ici / 2):
            skor += SETTINGS.EXTREME_WEEKEND_PENALTY

        if hafta_ici > 0 and hafta_sonu / hafta_ici >= 0.50:
            skor += SETTINGS.EXTREME_WEEKEND_PENALTY * 1.5

        if oran_farki < 0:
            skor -= min(8, abs(oran_farki) * 10)

    else:
        if oran_farki < 0:
            skor += SETTINGS.WEEKDAY_BALANCE_PENALTY * (abs(oran_farki) ** SETTINGS.BALANCE_POWER) * (toplam + 2)

        if adet_farki < 0:
            skor += SETTINGS.COUNT_GAP_PENALTY_WEEKDAY * (abs(adet_farki) ** 2)

        if toplam >= 5 and hafta_sonu == 0:
            skor += 25
        elif toplam >= 6 and hafta_sonu <= toplam * 0.12:
            skor += 18

    return skor


def score_monthly_coverage(p, d, monthly_assignment_counts, month_total_days):
    skor = 0
    ay_ici_adet = monthly_assignment_counts.get(p, 0)

    if ay_ici_adet == 0:
        skor -= SETTINGS.MONTHLY_ZERO_ASSIGNMENT_BONUS

        progress = month_progress_ratio(d, month_total_days)

        if progress >= SETTINGS.LATE_MONTH_RATIO:
            skor -= SETTINGS.LATE_MONTH_EXTRA_ZERO_BONUS

        if progress >= SETTINGS.VERY_LATE_MONTH_RATIO:
            skor -= SETTINGS.VERY_LATE_MONTH_EXTRA_ZERO_BONUS

    elif ay_ici_adet >= 2:
        skor += (ay_ici_adet - 1) * SETTINGS.MONTHLY_TWO_PLUS_PENALTY

    return skor


def score_recent_bayram_penalty(p, d, bayram_dates, arefe_dates, tatil, arefe):
    skor = 0
    tur = bayram_turu_of_date(d)

    if d in arefe or tur == "AREFE":
        gap = get_last_arefe_gap_days(arefe_dates, p, d)
        if gap <= SETTINGS.VERY_RECENT_BAYRAM_WINDOW_DAYS:
            skor += SETTINGS.AREFE_RECENT_PENALTY
        return skor

    if d in tatil:
        gap = get_last_bayram_gap_days(bayram_dates, p, d)

        if gap <= SETTINGS.RECENT_BAYRAM_WINDOW_DAYS:
            skor += SETTINGS.RECENT_BAYRAM_PENALTY

        if gap <= SETTINGS.VERY_RECENT_BAYRAM_WINDOW_DAYS:
            skor += SETTINGS.VERY_RECENT_BAYRAM_PENALTY

        same_type_recent = count_recent_bayram_same_type(
            bayram_dates=bayram_dates,
            p=p,
            d=d,
            target_type=tur,
            window_days=SETTINGS.RECENT_BAYRAM_WINDOW_DAYS,
        )
        skor += same_type_recent * SETTINGS.SAME_BAYRAM_TYPE_RECENT_PENALTY

    return skor


def score_person(
    p,
    d,
    w,
    tatil,
    arefe,
    totals,
    counts,
    weekday_stats,
    last_dates,
    monthly_assignment_counts,
    month_total_days,
    bayram_dates,
    arefe_dates,
):
    skor = totals[p] * SETTINGS.DENGE_KATSAYI + counts[p]

    if weekday_stats[p][d.weekday()] >= SETTINGS.MAX_SAME_WEEKDAY:
        skor += SETTINGS.WEEKDAY_PENALTY

    gap = get_gap_days(last_dates, p, d)
    skor -= min(gap, 30) * 0.05

    if gap > SETTINGS.MAX_GAP_DAYS:
        skor -= SETTINGS.MAX_GAP_BONUS
        skor -= (gap - SETTINGS.MAX_GAP_DAYS) * SETTINGS.MAX_GAP_EXTRA_PER_DAY

    if w > 1.4:
        skor += SETTINGS.AGIR_GUN_FRENI * weekday_stats[p][d.weekday()]

    skor += score_weekend_balance(p, d, weekday_stats)
    skor += score_monthly_coverage(p, d, monthly_assignment_counts, month_total_days)
    skor += score_recent_bayram_penalty(p, d, bayram_dates, arefe_dates, tatil, arefe)

    return skor + random.random() * 0.01


# =========================================================
# 104 ECZANELİ GRUPLAR
# =========================================================

def create_groups():
    groups = {
        "A1": ["ECZANE1", "ECZANE2", "ECZANE3", "ECZANE4", "ECZANE5", "ECZANE6", "ECZANE7", "ECZANE8", "ECZANE9"],
        "A2": ["ECZANE10", "ECZANE11", "ECZANE12", "ECZANE13", "ECZANE14", "ECZANE15", "ECZANE16", "ECZANE17", "ECZANE18"],
        "A3": ["ECZANE19", "ECZANE20", "ECZANE21", "ECZANE22", "ECZANE23", "ECZANE24", "ECZANE25", "ECZANE26", "ECZANE27"],

        "B1": ["ECZANE28", "ECZANE29", "ECZANE30", "ECZANE31", "ECZANE32", "ECZANE33", "ECZANE34", "ECZANE35"],
        "B2": ["ECZANE36", "ECZANE37", "ECZANE38", "ECZANE39", "ECZANE40", "ECZANE41", "ECZANE42", "ECZANE43", "ECZANE44"],
        "B3": ["ECZANE45", "ECZANE46", "ECZANE47", "ECZANE48", "ECZANE49", "ECZANE50", "ECZANE51", "ECZANE52"],

        "C1": ["ECZANE53", "ECZANE54", "ECZANE55", "ECZANE56", "ECZANE57", "ECZANE58", "ECZANE59", "ECZANE60"],
        "C2": ["ECZANE61", "ECZANE62", "ECZANE63", "ECZANE64", "ECZANE65", "ECZANE66", "ECZANE67", "ECZANE68", "ECZANE69"],
        "C3": ["ECZANE70", "ECZANE71", "ECZANE72", "ECZANE73", "ECZANE74", "ECZANE75", "ECZANE76", "ECZANE77", "ECZANE78"],

        "D1": ["ECZANE79", "ECZANE80", "ECZANE81", "ECZANE82", "ECZANE83", "ECZANE84", "ECZANE85", "ECZANE86"],
        "D2": ["ECZANE87", "ECZANE88", "ECZANE89", "ECZANE90", "ECZANE91", "ECZANE92", "ECZANE93", "ECZANE94", "ECZANE95"],
        "D3": ["ECZANE96", "ECZANE97", "ECZANE98", "ECZANE99", "ECZANE100", "ECZANE101", "ECZANE102", "ECZANE103", "ECZANE104"],
    }

    for eczane, data in eklenme.items():
        grup = str(data["grup"]).upper().strip()
        eczane = normalize_name(eczane)
        if grup in groups and eczane not in groups[grup]:
            groups[grup].append(eczane)

    return groups


# =========================================================
# ROTASYON - İKİNCİ KOD MANTIĞI
# Günde 4 eczane seçilir.
# =========================================================

KOMB_ABC = [
    ("A1", "B2", "C3", "D1"),
    ("B1", "C2", "A3", "D3"),
    ("C1", "A2", "B3", "D2"),
    ("A1", "C2", "B3", "D1"),
    ("B1", "A2", "C3", "D2"),
    ("C1", "B2", "A3", "D3"),
]


# =========================================================
# SEÇİM
# =========================================================

def zorunlu_secim(
    grup,
    d,
    w,
    tatil,
    arefe,
    totals,
    counts,
    weekday_stats,
    last_dates,
    bayram_stats,
    bayram_year_stats,
    bayram_dates,
    arefe_dates,
    monthly_assignment_counts,
    month_total_days,
    today_used,
):
    grup_norm = [normalize_name(p) for p in grup]

    aktifler = []
    for p in grup_norm:
        if p in today_used:
            continue
        if not is_person_active_on_date(p, d):
            continue
        aktifler.append(p)

    if not aktifler:
        aktifler = [p for p in grup_norm if p not in today_used] or list(grup_norm)

    kademe1 = []
    kademe2 = []
    kademe3 = []

    for p in aktifler:
        gap = get_gap_days(last_dates, p, d)

        if gap >= SETTINGS.MIN_GAP_DAYS and weekday_stats[p][d.weekday()] < SETTINGS.MAX_SAME_WEEKDAY:
            kademe1.append(p)
        elif weekday_stats[p][d.weekday()] < SETTINGS.MAX_SAME_WEEKDAY:
            kademe2.append(p)
        else:
            kademe3.append(p)

    adaylar = kademe1 or kademe2 or kademe3 or list(aktifler)

    max_gap_asanlar = [p for p in adaylar if get_gap_days(last_dates, p, d) > SETTINGS.MAX_GAP_DAYS]
    if max_gap_asanlar:
        adaylar = max_gap_asanlar

    if d in tatil:
        # 1) Bayram yasaklı eczaneler elenir.
        adaylar = [p for p in adaylar if p not in BAYRAM_YASAK]

        # 2) Aynı yıl ikinci kez bayram tamamen engellenir.
        bayram_uygun = [p for p in adaylar if bayram_year_stats[p][d.year] == 0]
        if bayram_uygun:
            adaylar = bayram_uygun

        # 3) Toplam bayram geçmişi az olan öne alınır.
        if adaylar:
            min_b = min(bayram_stats[p] for p in adaylar)
            adaylar = [p for p in adaylar if bayram_stats[p] == min_b]

    elif d in arefe:
        # Arefede de geçmiş arefesi az olanlar öne alınır.
        if adaylar:
            min_arefe = min(len(arefe_dates[p]) for p in adaylar)
            adaylar = [p for p in adaylar if len(arefe_dates[p]) == min_arefe]

    if not adaylar:
        if d in tatil:
            adaylar = [
                p for p in aktifler
                if p not in BAYRAM_YASAK and bayram_year_stats[p][d.year] == 0
            ]

    if not adaylar:
        adaylar = [p for p in aktifler if p not in today_used] or list(aktifler) or list(grup_norm)

    if not adaylar:
        raise Exception(f"Aday bulunamadı | Tarih={d} | Grup={grup}")

    pick = min(
        adaylar,
        key=lambda p: score_person(
            p=p,
            d=d,
            w=w,
            tatil=tatil,
            arefe=arefe,
            totals=totals,
            counts=counts,
            weekday_stats=weekday_stats,
            last_dates=last_dates,
            monthly_assignment_counts=monthly_assignment_counts,
            month_total_days=month_total_days,
            bayram_dates=bayram_dates,
            arefe_dates=arefe_dates,
        ),
    )

    return pick


# =========================================================
# AYLIK PLAN
# =========================================================

def generate_month(
    groups,
    year,
    month,
    totals,
    counts,
    weekday_stats,
    bayram_stats,
    last_dates,
    bayram_year_stats,
    bayram_dates,
    arefe_dates,
):
    tatil = turkiye_tatilleri(year)
    arefe = arefe_gunleri(year)

    first = date(year, month, 1)
    dim = calendar.monthrange(year, month)[1]

    schedule = {}
    monthly_assignment_counts = defaultdict(int)

    for i in range(dim):
        d = first + timedelta(days=i)
        w = day_weight(d, tatil, arefe)

        picks = {}
        today_used = set()

        for g in KOMB_ABC[i % 6]:
            pick = zorunlu_secim(
                grup=groups[g],
                d=d,
                w=w,
                tatil=tatil,
                arefe=arefe,
                totals=totals,
                counts=counts,
                weekday_stats=weekday_stats,
                last_dates=last_dates,
                bayram_stats=bayram_stats,
                bayram_year_stats=bayram_year_stats,
                bayram_dates=bayram_dates,
                arefe_dates=arefe_dates,
                monthly_assignment_counts=monthly_assignment_counts,
                month_total_days=dim,
                today_used=today_used,
            )

            picks[g] = pick
            today_used.add(pick)

            totals[pick] += w
            counts[pick] += 1
            weekday_stats[pick][d.weekday()] += 1
            last_dates[pick] = d
            monthly_assignment_counts[pick] += 1

            key = (d.year, d.month)

            if d in tatil:
                bayram_stats[pick] += 1
                bayram_year_stats[pick][d.year] += 1
                bayram_dates[pick].append(d)
                monthly_stats[pick][key]["bayram"] += 1
            elif d in arefe:
                arefe_dates[pick].append(d)
                monthly_stats[pick][key]["arefe"] += 1
            else:
                monthly_stats[pick][key][DAY_KEYS[d.weekday()]] += 1

        schedule[d] = picks

    return schedule


# =========================================================
# ANA MOTOR
# =========================================================

def main(y, m, nm, GECMIS_YUK, GECMIS_BAYRAM=None, output_name="Alternatif.xlsx"):
    global monthly_stats

    monthly_stats = defaultdict(lambda: defaultdict(lambda: {
        "bayram": 0,
        "arefe": 0,
        "pazartesi": 0,
        "salı": 0,
        "çarşamba": 0,
        "perşembe": 0,
        "cuma": 0,
        "cumartesi": 0,
        "pazar": 0,
    }))

    if GECMIS_BAYRAM is None:
        GECMIS_BAYRAM = []

    groups = create_groups()

    totals = {normalize_name(p): 0 for g in groups.values() for p in g}
    counts = {normalize_name(p): 0 for g in groups.values() for p in g}
    weekday_stats = {normalize_name(p): {i: 0 for i in range(7)} for p in totals}
    bayram_stats = {normalize_name(p): 0 for p in totals}
    bayram_year_stats = defaultdict(lambda: defaultdict(int))
    bayram_dates = defaultdict(list)
    arefe_dates = defaultdict(list)
    last_dates = {}
    gecmis_katsayi_map = {}

    for p, v in GECMIS_YUK.items():
        p = normalize_name(p)

        if p not in totals:
            continue

        kats = (
            v.get("pazartesi", 0) +
            v.get("salı", 0) +
            v.get("çarşamba", 0) +
            v.get("perşembe", 0) +
            v.get("cuma", 0) +
            v.get("cumartesi", 0) * 1.5 +
            v.get("pazar", 0) * 2
        )

        totals[p] += kats
        counts[p] += (
            v.get("pazartesi", 0) +
            v.get("salı", 0) +
            v.get("çarşamba", 0) +
            v.get("perşembe", 0) +
            v.get("cuma", 0) +
            v.get("cumartesi", 0) +
            v.get("pazar", 0) +
            v.get("bayram", 0)
        )

        bayram_stats[p] += v.get("bayram", 0)

        weekday_stats[p][0] += v.get("pazartesi", 0)
        weekday_stats[p][1] += v.get("salı", 0)
        weekday_stats[p][2] += v.get("çarşamba", 0)
        weekday_stats[p][3] += v.get("perşembe", 0)
        weekday_stats[p][4] += v.get("cuma", 0)
        weekday_stats[p][5] += v.get("cumartesi", 0)
        weekday_stats[p][6] += v.get("pazar", 0)

        gecmis_katsayi_map[p] = round(kats, 2)

    for p in totals:
        gecmis_katsayi_map.setdefault(p, 0)

    for rec in GECMIS_BAYRAM:
        p = normalize_name(rec["eczane"])
        d = rec["tarih"]
        tur = rec["tur"]

        if p not in totals:
            continue

        if tur == "AREFE":
            arefe_dates[p].append(d)
        else:
            bayram_dates[p].append(d)
            bayram_year_stats[p][d.year] += 1
            bayram_stats[p] = max(bayram_stats[p], len(bayram_dates[p]))

        if p not in last_dates or d > last_dates[p]:
            last_dates[p] = d

    wb = Workbook()
    header = ["Tarih", "Gün"] + list(groups.keys())

    for k in range(nm):
        year = y + ((m - 1 + k) // 12)
        month = ((m - 1 + k) % 12) + 1

        ws = wb.create_sheet(f"{year}-{month:02d}")
        ws.append(header)

        sched = generate_month(
            groups=groups,
            year=year,
            month=month,
            totals=totals,
            counts=counts,
            weekday_stats=weekday_stats,
            bayram_stats=bayram_stats,
            last_dates=last_dates,
            bayram_year_stats=bayram_year_stats,
            bayram_dates=bayram_dates,
            arefe_dates=arefe_dates,
        )

        for d, p in sorted(sched.items()):
            row = [d.strftime("%d.%m.%Y"), DAY_SHORT[d.weekday()]]
            for g in groups:
                row.append(p.get(g, ""))
            ws.append(row)

        for c in ws[1]:
            c.font = Font(bold=True)

    summary = wb.create_sheet("GENEL OZET")
    summary.append([
        "Eczane",
        "Grup",
        "Geçmiş Katsayı",
        "Geçmiş Bayram",
        "Toplam Nöbet",
        "Toplam Katsayı",
        "Bayram",
        "Pzt", "Salı", "Çarş", "Perş", "Cuma", "Ctesi", "Pazar",
    ])

    eczane_grup = {normalize_name(p): g for g, plist in groups.items() for p in plist}

    for p in totals:
        toplam_katsayi = (
            weekday_stats[p][0] +
            weekday_stats[p][1] +
            weekday_stats[p][2] +
            weekday_stats[p][3] +
            weekday_stats[p][4] +
            weekday_stats[p][5] * 1.5 +
            weekday_stats[p][6] * 2
        )

        summary.append([
            p,
            eczane_grup.get(p, ""),
            gecmis_katsayi_map.get(p, 0),
            GECMIS_YUK.get(p, {}).get("bayram", 0),
            counts[p],
            round(toplam_katsayi, 2),
            bayram_stats[p],
            weekday_stats[p][0],
            weekday_stats[p][1],
            weekday_stats[p][2],
            weekday_stats[p][3],
            weekday_stats[p][4],
            weekday_stats[p][5],
            weekday_stats[p][6],
        ])

    for c in summary[1]:
        c.font = Font(bold=True)

    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    wb.save(output_name)

    wb2 = Workbook()
    ws2 = wb2.active
    ws2.title = "AYLIK DETAY"

    ws2.append([
        "Eczane",
        "Yıl",
        "Ay",
        "Bayram",
        "Arefe",
        "Pazartesi",
        "Salı",
        "Çarşamba",
        "Perşembe",
        "Cuma",
        "Cumartesi",
        "Pazar",
    ])

    for eczane in sorted(totals.keys()):
        if eczane not in monthly_stats:
            continue

        for (yil, ay), veri in sorted(monthly_stats[eczane].items()):
            ws2.append([
                eczane,
                yil,
                ay,
                veri["bayram"],
                veri["arefe"],
                veri["pazartesi"],
                veri["salı"],
                veri["çarşamba"],
                veri["perşembe"],
                veri["cuma"],
                veri["cumartesi"],
                veri["pazar"],
            ])

    for c in ws2[1]:
        c.font = Font(bold=True)

    detail_name = "aylik_nobet_data.xlsx"
    wb2.save(detail_name)

    return output_name, detail_name


# =========================================================
# STREAMLIT / WEB ÇAĞIRMA
# =========================================================

def run_schedule(
    y,
    m,
    nm,
    gecmis_yuk_df,
    gecmis_bayram_df=None,
    eklenme_input=None,
    cikma_input=None,
):
    global eklenme_tarihi
    global cikma_tarihi
    global eklenme
    global monthly_stats

    if eklenme_input is None:
        eklenme_input = {}
    if cikma_input is None:
        cikma_input = {}

    monthly_stats = defaultdict(lambda: defaultdict(lambda: {
        "bayram": 0,
        "arefe": 0,
        "pazartesi": 0,
        "salı": 0,
        "çarşamba": 0,
        "perşembe": 0,
        "cuma": 0,
        "cumartesi": 0,
        "pazar": 0,
    }))

    eklenme = {
        normalize_name(k): {
            "grup": str(v["grup"]).upper().strip(),
            "tarih": v["tarih"],
        }
        for k, v in eklenme_input.items()
    }
    eklenme_tarihi = {normalize_name(k): v["tarih"] for k, v in eklenme_input.items()}
    cikma_tarihi = {normalize_name(k): v for k, v in cikma_input.items()}

    GECMIS_YUK = excelden_gecmis_yuk_df(gecmis_yuk_df)
    GECMIS_BAYRAM = excelden_gecmis_bayram_df(gecmis_bayram_df) if gecmis_bayram_df is not None else []

    return main(y, m, nm, GECMIS_YUK, GECMIS_BAYRAM, output_name="Alternatif.xlsx")


# =========================================================
# MASAÜSTÜ / KONSOL ÇALIŞTIRMA
# =========================================================

if __name__ == "__main__":
    GECMIS_YUK = excelden_gecmis_yuk("gecmis_nobet_erzurum.xlsx")

    y = int(input("Yıl: "))
    m = int(input("Başlangıç ayı: "))
    nm = int(input("Kaç ay?: "))

    output, detail = main(y, m, nm, GECMIS_YUK, GECMIS_BAYRAM=[])
    print(f"✅ Plan hazır: {output}")
    print(f"📊 Aylık detay hazır: {detail}")
