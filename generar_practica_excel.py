#!/usr/bin/env python3
"""
Generador de práctica de Excel para la vacante
"Analyst, Customer Success Operations" (Mitratech, remoto México).

Crea dos libros de Excel con datos ficticios (pero realistas) de una empresa
SaaS B2B y ejercicios de lo que pide el puesto:

  * CS_Ops_Practica.xlsx   -> ejercicios con celdas amarillas por llenar y
                              columnas de verificación (✓ / ✗) automáticas.
  * CS_Ops_Soluciones.xlsx -> el mismo libro con todas las fórmulas resueltas.
  * csv/                   -> las tablas en CSV para practicar Power Query.

Temas: XLOOKUP / VLOOKUP / INDEX-MATCH, SUMIFS / COUNTIFS / AVERAGEIFS /
MAXIFS / MINIFS, SUMPRODUCT, IFS / SWITCH, limpieza de datos (TRIM, PROPER,
NUMBERVALUE, fechas en texto), duplicados, health score de cuentas, auditoría
de uso de reportes, tablas dinámicas, funciones de matriz dinámica (FILTER,
UNIQUE, SORT, LET, TAKE...), tendencias, churn, renovaciones y un dashboard
para CSMs.

Uso:
    pip install xlsxwriter
    python generar_practica_excel.py                  # datos con semilla 42
    python generar_practica_excel.py --semilla 7      # otro set de datos
    python generar_practica_excel.py --cuentas 250 --salida mis_archivos

Cada semilla genera datos distintos, así puedes repetir los ejercicios con
números nuevos. Las fórmulas funcionan en Excel 365 / Excel 2021+ (las de
matriz dinámica como LET, TAKE y HSTACK requieren Excel 365).
"""

import argparse
import csv
import datetime as dt
import math
import os
import random
import re
import unicodedata
from collections import Counter, defaultdict

try:
    import xlsxwriter
    from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell
except ImportError:  # pragma: no cover
    raise SystemExit("Falta la librería xlsxwriter. Instálala con:  pip install xlsxwriter")


# ---------------------------------------------------------------------------
# Utilidades de fechas y de "imitación" de funciones de Excel
# ---------------------------------------------------------------------------

EPOCH = dt.date(1899, 12, 30)


def serial(d):
    """Fecha -> número de serie de Excel."""
    return (d - EPOCH).days


def month_add(d, n):
    y = d.year + (d.month - 1 + n) // 12
    m = (d.month - 1 + n) % 12 + 1
    return dt.date(y, m, 1)


def eomonth(d):
    return month_add(d, 1) - dt.timedelta(days=1)


def edate(d, n):
    first = month_add(d, n)
    return first.replace(day=min(d.day, eomonth(first).day))


def excel_round(x, digits=0):
    """ROUND de Excel (mitad hacia arriba, alejándose de cero)."""
    m = 10 ** digits
    v = math.floor(abs(x) * m + 0.5 + 1e-9) / m
    return v if x >= 0 else -v


def excel_trim(s):
    return " ".join(p for p in s.split(" ") if p)


def excel_proper(s):
    out, prev_alpha = [], False
    for ch in s:
        if ch.isalpha():
            out.append(ch.lower() if prev_alpha else ch.upper())
            prev_alpha = True
        else:
            out.append(ch)
            prev_alpha = False
    return "".join(out)


def ascii_slug(s):
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().lower()


def clamp(x, lo, hi):
    return max(lo, min(hi, x))


# ---------------------------------------------------------------------------
# Datos ficticios
# ---------------------------------------------------------------------------

PREFIXES = ["Northwind", "Blue Ridge", "Summit", "Atlas", "Pioneer", "Evergreen", "Crescent",
            "Ironwood", "Silverline", "Harbor", "Redwood", "Keystone", "Meridian", "Lakeshore",
            "Granite", "Aurora", "Sterling", "Cobalt", "Vantage", "Horizon", "Brightpath",
            "Oakmont", "Riverbend", "Westfield", "Pinnacle", "Falcon", "Monarch", "Sierra",
            "Coastal", "Liberty", "Anchor", "Bluewater", "Cedar", "Driftwood", "Elmstone"]
SUFFIXES = ["Legal", "Insurance", "Health", "Bank", "Energy", "Manufacturing", "Logistics",
            "Partners", "Group", "Holdings", "Financial", "Pharma", "Retail", "Capital", "Systems"]

SEGMENTS = ["Enterprise", "Mid-Market", "SMB"]
REGIONS = ["North America", "LATAM", "EMEA", "APAC"]
INDUSTRIES = ["Legal Services", "Financial Services", "Healthcare", "Manufacturing",
              "Insurance", "Technology", "Public Sector", "Energy"]
PRODUCTS = ["Legal Spend Mgmt", "Matter Management", "Contract Lifecycle", "GRC / Risk",
            "HR Compliance", "Workflow Automation"]
TICKET_CATEGORIES = ["Bug", "How-to", "Integration", "Performance", "Billing",
                     "Feature Request", "Data/Reporting"]
PRIORITIES = ["P1", "P2", "P3", "P4"]
ROLES = ["CSM", "Account Manager", "CS Leadership", "Executive", "Support", "Product",
         "CS Ops", "Partner Mgmt"]
TIERS = [(0, "Bronze"), (25000, "Silver"), (100000, "Gold"), (250000, "Platinum")]
DOMAIN = "empresa-demo.com"

CSM_SEED = [
    ("Ana López", "Enterprise", "Laura Kim"),
    ("Carlos Méndez", "Enterprise", "Laura Kim"),
    ("Sofía Ramírez", "Mid-Market", "Laura Kim"),
    ("Daniel Ortiz", "Mid-Market", "Robert Hayes"),
    ("Emily Carter", "Mid-Market", "Robert Hayes"),
    ("James Walker", "SMB", "Robert Hayes"),
    ("Lucía Herrera", "SMB", "Robert Hayes"),
    ("Michael Brooks", "SMB", "Laura Kim"),
]

# (nombre, tipo, audiencia, frecuencia, fuente, perfil de uso)
REPORT_SEED = [
    ("Customer Health Scorecard", "Standard", "CS Leadership", "Weekly", "Gainsight", "hot"),
    ("Weekly Active Users by Account", "Standard", "CSM", "Weekly", "Product DB", "hot"),
    ("Renewal Pipeline - Next 120 Days", "Standard", "CS Leadership", "Weekly", "Salesforce", "hot"),
    ("QBR Usage Pack - Enterprise", "Custom", "CSM", "Quarterly", "Product DB", "warm"),
    ("Seat Utilization by Product", "Standard", "CSM", "Monthly", "Product DB", "warm"),
    ("Ticket Backlog by CSM", "Standard", "CSM", "Daily", "Salesforce", "hot"),
    ("Executive ARR at Risk", "Custom", "Executive", "Monthly", "Salesforce", "warm"),
    ("Login Report (Legacy v1)", "Standard", "CSM", "Weekly", "Product DB", "dead"),
    ("Login Report v2", "Standard", "CSM", "Weekly", "Product DB", "warm"),
    ("Feature Adoption Heatmap", "Custom", "Product", "Monthly", "Product DB", "cold"),
    ("Contractual Usage Report - Monthly", "Contractual", "Customer", "Monthly", "Product DB", "warm"),
    ("Contractual SLA Report - Quarterly", "Contractual", "Customer", "Quarterly", "Salesforce", "cold"),
    ("NPS Verbatims Digest", "Standard", "CS Leadership", "Monthly", "Gainsight", "warm"),
    ("Churn Post-Mortem Tracker", "Custom", "CS Leadership", "Ad hoc", "Salesforce", "stale"),
    ("CSM Book of Business", "Standard", "CSM", "Weekly", "Salesforce", "hot"),
    ("Onboarding Progress (Implementation)", "Standard", "CSM", "Weekly", "Gainsight", "warm"),
    ("Support CSAT Trend", "Standard", "Support", "Monthly", "Salesforce", "cold"),
    ("Ad hoc - Partner Usage Extract", "Custom", "Partner", "Ad hoc", "Product DB", "hot"),
    ("Ad hoc - Upsell Whitespace", "Custom", "Account Mgmt", "Ad hoc", "Salesforce", "hot"),
    ("Monthly Logins Summary (Old)", "Standard", "CS Leadership", "Monthly", "Product DB", "dead"),
    ("Copy of Customer Health Scorecard", "Custom", "CSM", "Weekly", "Gainsight", "stale"),
    ("Product Release Adoption", "Custom", "Product", "Monthly", "Product DB", "cold"),
    ("Integration Errors by Account", "Standard", "Support", "Daily", "Product DB", "warm"),
    ("License Overage Report", "Standard", "Account Mgmt", "Monthly", "Salesforce", "warm"),
    ("EBR Deck Data Pull", "Custom", "CS Leadership", "Quarterly", "Gainsight", "cold"),
    ("Inactive Users Alert", "Standard", "CSM", "Weekly", "Product DB", "warm"),
    ("Report Usage Audit (Meta)", "Custom", "CS Ops", "Monthly", "Product DB", "cold"),
    ("Regional ARR Summary - LATAM", "Custom", "CS Leadership", "Monthly", "Salesforce", "stale"),
    ("Regional ARR Summary - EMEA", "Custom", "CS Leadership", "Monthly", "Salesforce", "dead"),
    ("Customer Success Plan Status", "Standard", "CSM", "Monthly", "Gainsight", "cold"),
    ("Contractual Usage Report - Healthcare", "Contractual", "Customer", "Monthly", "Product DB", "dead"),
    ("Ad hoc - Price Increase Impact", "Custom", "Executive", "Ad hoc", "Salesforce", "stale"),
    ("Time to First Value", "Standard", "CS Leadership", "Monthly", "Gainsight", "warm"),
    ("Escalations Log", "Standard", "CS Leadership", "Weekly", "Salesforce", "warm"),
    ("Daily Usage Firehose", "Standard", "CS Ops", "Daily", "Product DB", "cold"),
    ("Partner Health Overview", "Custom", "Partner", "Monthly", "Gainsight", "cold"),
]

AUDIENCE_ROLES = {
    "CSM": [("CSM", 0.8), ("CS Leadership", 0.1), ("Account Manager", 0.1)],
    "CS Leadership": [("CS Leadership", 0.6), ("CSM", 0.2), ("Executive", 0.2)],
    "Executive": [("Executive", 0.7), ("CS Leadership", 0.3)],
    "Customer": [("CSM", 0.9), ("CS Ops", 0.1)],
    "Product": [("Product", 0.8), ("CS Ops", 0.2)],
    "Support": [("Support", 0.8), ("CSM", 0.2)],
    "Account Mgmt": [("Account Manager", 0.85), ("CS Leadership", 0.15)],
    "Partner": [("Partner Mgmt", 0.9), ("CSM", 0.1)],
    "CS Ops": [("CS Ops", 1.0)],
}

EXTRA_USERS = [
    ("Account Manager", ["Valeria Cruz", "Tom Becker", "Andrea Navarro", "Chris Doyle", "Fernanda Ruiz"]),
    ("CS Leadership", ["Laura Kim", "Robert Hayes", "Patricia Gómez"]),
    ("Executive", ["Steven Grant", "Mónica Salinas"]),
    ("Support", ["Jorge Castillo", "Hannah Lee", "Diego Vargas", "Nora Ellis", "Iván Pérez"]),
    ("Product", ["Priya Shah", "Alejandro Ríos", "Megan Ford"]),
    ("CS Ops", ["Mariana Torres", "Kevin Liu"]),
    ("Partner Mgmt", ["Ricardo Peña", "Olivia Grant"]),
]
REPORT_OWNERS = ["Mariana Torres", "Kevin Liu", "Paola Sánchez", "Ryan Mitchell"]


class Data:
    pass


def generate(seed, n_accounts, as_of):
    rng = random.Random(seed)
    D = Data()
    D.as_of = as_of
    D.last_month = as_of.replace(day=1)
    D.months = [month_add(D.last_month, -i) for i in range(11, -1, -1)]
    D.future_months = [month_add(D.last_month, i) for i in range(1, 7)]

    # --- CSMs
    D.csms = []
    for i, (name, team, mgr) in enumerate(CSM_SEED):
        D.csms.append(dict(CSM_ID=f"CSM-{i + 1:02d}", CSM_Name=name,
                           Email=f"{ascii_slug(name).replace(' ', '.')}@{DOMAIN}",
                           Team=team, Manager=mgr))
    csm_by_team = defaultdict(list)
    for c in D.csms:
        csm_by_team[c["Team"]].append(c["CSM_ID"])

    # --- Cuentas
    names = rng.sample([f"{p} {s}" for p in PREFIXES for s in SUFFIXES], n_accounts)
    D.accounts = []
    for i in range(n_accounts):
        seg = rng.choices(SEGMENTS, [0.2, 0.4, 0.4])[0]
        arr_rng, seat_rng = {"Enterprise": ((150000, 600000), (200, 1500)),
                             "Mid-Market": ((40000, 150000), (50, 300)),
                             "SMB": ((8000, 40000), (10, 60))}[seg]
        h = rng.betavariate(3, 2)
        churned = h < 0.4 and rng.random() < 0.45
        start = as_of - dt.timedelta(days=rng.randint(400, 1800))
        if churned:
            churn_date = as_of - dt.timedelta(days=rng.randint(5, 330))
            renewal = churn_date
        else:
            churn_date = None
            renewal = as_of + dt.timedelta(days=rng.randint(3, 365))
        if seg == "Enterprise":
            rep = rng.choices(["Monthly", "Quarterly", "None"], [0.5, 0.3, 0.2])[0]
        elif seg == "Mid-Market":
            rep = rng.choices(["Quarterly", "None"], [0.2, 0.8])[0]
        else:
            rep = "None"
        D.accounts.append(dict(
            AccountID=f"ACC-{i + 1:04d}", AccountName=names[i], Segment=seg,
            Region=rng.choices(REGIONS, [0.45, 0.25, 0.2, 0.1])[0],
            Industry=rng.choice(INDUSTRIES), Product=rng.choice(PRODUCTS),
            CSM_ID=rng.choice(csm_by_team[seg]),
            ARR_USD=round(rng.uniform(*arr_rng) / 100) * 100,
            LicensedSeats=rng.randint(*seat_rng),
            ContractStart=start, RenewalDate=renewal,
            Status="Churned" if churned else "Active", ChurnDate=churn_date,
            ContractualReporting=rep,
            _h=h, _slope=rng.uniform(-0.035, -0.01) if h < 0.45 else rng.uniform(-0.01, 0.02)))
    D.acc_by_id = {a["AccountID"]: a for a in D.accounts}

    # --- Uso mensual (ordenado por mes y luego por cuenta, como un export típico)
    D.usage = []
    for i, m in enumerate(D.months):
        for a in D.accounts:
            if a["ChurnDate"] and m > a["ChurnDate"]:
                continue
            h = a["_h"]
            util = clamp(0.12 + 0.8 * h + a["_slope"] * (i - 5.5) + rng.gauss(0, 0.04), 0.03, 1.0)
            au = max(1, round(a["LicensedSeats"] * util))
            D.usage.append(dict(
                AccountID=a["AccountID"], Month=m, ActiveUsers=au,
                Logins=round(au * rng.uniform(5, 18)),
                ReportsRun=round(au * rng.uniform(0.3, 2.5)),
                FeaturesAdopted=int(clamp(round(1 + 8 * h + rng.gauss(0, 0.8)), 1, 10)),
                AvgSessionMin=round(rng.uniform(5, 40), 1)))

    # --- Tickets de soporte
    tickets = []
    lower = as_of - dt.timedelta(days=364)
    for a in D.accounts:
        h = a["_h"]
        factor = {"Enterprise": 2.0, "Mid-Market": 1.3, "SMB": 0.8}[a["Segment"]]
        n = min(25, int(rng.expovariate(1 / ((2 + 8 * (1 - h)) * factor))))
        upper = a["ChurnDate"] or as_of
        if upper < lower:
            continue
        for _ in range(n):
            opened = lower + dt.timedelta(days=rng.randint(0, (upper - lower).days))
            pr = rng.choices(PRIORITIES, [0.03 + 0.1 * (1 - h), 0.12 + 0.2 * (1 - h), 0.45, 0.3])[0]
            res = int(rng.expovariate(1 / {"P1": 2, "P2": 5, "P3": 9, "P4": 14}[pr]))
            closed = opened + dt.timedelta(days=res)
            if closed > as_of:
                closed = None
            csat = None
            if closed and rng.random() < 0.55:
                csat = int(clamp(round(rng.gauss(2 + 3 * h, 0.9)), 1, 5))
            tickets.append(dict(AccountID=a["AccountID"], OpenedDate=opened, ClosedDate=closed,
                                Priority=pr, Category=rng.choice(TICKET_CATEGORIES), CSAT=csat))
    tickets.sort(key=lambda t: (t["OpenedDate"], t["AccountID"]))
    D.tickets = [dict(TicketID=f"TCK-{i + 1:05d}", **t) for i, t in enumerate(tickets)]

    # --- Encuestas NPS (algunas con más de 12 meses de antigüedad, a propósito)
    nps = []
    for a in D.accounts:
        h = a["_h"]
        upper = a["ChurnDate"] or as_of
        span = 480 - (as_of - upper).days
        if span <= 0:
            continue
        for _ in range(rng.choices([0, 1, 2, 3], [0.25, 0.35, 0.25, 0.15])[0]):
            d = upper - dt.timedelta(days=rng.randint(0, span))
            nps.append(dict(AccountID=a["AccountID"], ResponseDate=d,
                            Score=int(clamp(round(rng.gauss(2.5 + 7.5 * h, 1.6)), 0, 10))))
    nps.sort(key=lambda r: (r["ResponseDate"], r["AccountID"]))
    D.nps = [dict(SurveyID=f"NPS-{i + 1:04d}", **r) for i, r in enumerate(nps)]

    # --- Catálogo de reportes
    D.reports = []
    for i, (name, typ, aud, freq, src, prof) in enumerate(REPORT_SEED):
        created = as_of - dt.timedelta(days=rng.randint(300, 1700))
        modified = created + dt.timedelta(days=rng.randint(0, (as_of - created).days))
        D.reports.append(dict(
            ReportID=f"RPT-{i + 1:03d}", ReportName=name, Type=typ, Audience=aud,
            Frequency=freq, Source=src, Owner=rng.choice(REPORT_OWNERS),
            CreatedDate=created, LastModified=modified,
            Recipients=rng.randint(1, 6) if aud in ("Executive", "Customer") else rng.randint(3, 40),
            MaintHrsPerMonth=rng.choice([0.5, 1, 1.5, 2, 3, 4, 6, 8]), _profile=prof))

    # --- Usuarios internos que consultan reportes
    users = [(c["Email"], "CSM") for c in D.csms]
    for role, people in EXTRA_USERS:
        users += [(f"{ascii_slug(p).replace(' ', '.')}@{DOMAIN}", role) for p in people]
    by_role = defaultdict(list)
    for email, role in users:
        by_role[role].append(email)
    user_w = {email: rng.uniform(0.2, 3) for email, _ in users}

    # --- Log de uso de reportes (últimos 180 días)
    log = []
    start = as_of - dt.timedelta(days=179)
    for r in D.reports:
        prof = r["_profile"]
        n = {"hot": rng.randint(250, 400), "warm": rng.randint(60, 140), "cold": rng.randint(6, 22),
             "stale": rng.randint(10, 30), "dead": 0}[prof]
        last_day = 79 if prof == "stale" else 179  # "stale": nadie lo abre en los últimos 100 días
        roles, weights = zip(*AUDIENCE_ROLES[r["Audience"]])
        for _ in range(n):
            role = rng.choices(roles, weights)[0]
            pool = by_role[role]
            email = rng.choices(pool, [user_w[u] for u in pool])[0]
            action = rng.choices(["View", "Export", "Subscribe"], [0.75, 0.2, 0.05])[0]
            dur = {"View": rng.randint(20, 600), "Export": rng.randint(5, 30),
                   "Subscribe": rng.randint(3, 10)}[action]
            log.append(dict(Date=start + dt.timedelta(days=rng.randint(0, last_day)),
                            ReportID=r["ReportID"], UserEmail=email, UserRole=role,
                            Action=action, DurationSec=dur))
    log.sort(key=lambda x: (x["Date"], x["ReportID"], x["UserEmail"]))
    D.log = [dict(LogID=f"LOG-{i + 1:05d}", **x) for i, x in enumerate(log)]

    # --- Export "sucio" del CRM para el ejercicio de limpieza
    D.raw = make_dirty_export(rng, D)
    D.active = sorted([a for a in D.accounts if a["Status"] == "Active"], key=lambda a: a["AccountID"])
    D.rng = rng
    return D


def make_dirty_export(rng, D):
    sample = rng.sample(D.accounts, min(48, len(D.accounts)))
    csm_email = {c["CSM_ID"]: c["Email"] for c in D.csms}
    rows = []

    def dirty(a, arr_override=None):
        id_ = a["AccountID"]
        id_ = rng.choice([id_, id_.lower(), f" {id_}", f"{id_}  ", id_])
        name = a["AccountName"]
        name = rng.choice([name, name.upper(), name.lower(), f"  {name}", name.replace(" ", "  ")])
        seg = {"Enterprise": ["Enterprise", "enterprise", "ENTERPRISE ", "Ent"],
               "Mid-Market": ["Mid-Market", "mid-market", " Mid Market", "MM"],
               "SMB": ["SMB", "smb", "Small Business", " SMB "]}[a["Segment"]]
        seg = rng.choice(seg)
        arr = arr_override if arr_override is not None else a["ARR_USD"]
        arr_txt = rng.choice([f"${arr:,.0f}", f"{arr}", f"{arr:,.0f} USD", f"USD {arr}",
                              f" {arr:.2f}", f"${arr:,.2f}"])
        d = a["RenewalDate"]
        ren = rng.choice([d.strftime("%Y-%m-%d"), d.strftime("%d/%m/%Y")])
        em = csm_email[a["CSM_ID"]]
        em = rng.choice([em, em.upper(), f"  {em}", em.title(), em, ""]) if rng.random() < 0.9 else ""
        return [id_, name, seg, arr_txt, ren, em]

    mismatch = set(rng.sample(range(len(sample)), 6))
    for i, a in enumerate(sample):
        over = a["ARR_USD"] + rng.choice([-5000, 2500, 10000, -1200]) if i in mismatch else None
        rows.append(dirty(a, over))
    for a in rng.sample(sample, 6):  # duplicados (mismo ID, distinto formato)
        rows.append(dirty(a))
    n = len(D.accounts)
    for k in range(4):  # cuentas que no existen en el maestro
        fake = dict(AccountID=f"ACC-{n + 101 + k:04d}", AccountName=rng.choice(PREFIXES) + " Ventures",
                    Segment=rng.choice(SEGMENTS), ARR_USD=rng.randint(10, 90) * 1000,
                    RenewalDate=D.as_of + dt.timedelta(days=rng.randint(10, 300)),
                    CSM_ID=rng.choice(D.csms)["CSM_ID"])
        rows.append(dirty(fake))
    rng.shuffle(rows)
    return rows


# ---------------------------------------------------------------------------
# Constructor del libro
# ---------------------------------------------------------------------------

NF_USD = "$#,##0"
NF_USD2 = "$#,##0.00"
NF_DATE = "yyyy-mm-dd"
NF_MONTH = "mmm yyyy"
NF_PCT = "0.0%"
NF_INT = "#,##0"
NF_DEC = "0.0"
NF_DEC2 = "0.00"

C_TITLE = "#1F3864"
C_INPUT = "#FFF2CC"
C_SOL = "#E2EFDA"
C_HINT = "#F2F2F2"
C_HDR = "#305496"

NOCHECK = object()
DYNAMIC = re.compile(r"\b(FILTER|UNIQUE|SORT|SORTBY|SEQUENCE|XLOOKUP|XMATCH|TAKE|DROP|CHOOSECOLS|"
                     r"HSTACK|VSTACK|LET)\(")

SHEETS = ["LEEME", "Diccionario", "E1_Lookups", "E2_Agregaciones", "E3_Health_Score",
          "E4_Auditoria_Reportes", "E5_Limpieza_Datos", "E6_Pivots", "E7_Dinamicas_365",
          "E8_Tendencias", "E9_Dashboard_CSM", "E10_PowerQuery", "Entrevista",
          "Accounts", "CSMs", "Usage_Monthly", "Support_Tickets", "NPS", "Report_Catalog",
          "Report_Usage_Log", "_key"]
EXERCISES = SHEETS[2:11]

# Tablas de Excel (ListObjects)
T_ACC, T_CSM, T_USE, T_TCK, T_NPS, T_REP, T_LOG = ("tbl_Accounts", "tbl_CSMs", "tbl_Usage",
                                                   "tbl_Tickets", "tbl_NPS", "tbl_Reports",
                                                   "tbl_ReportLog")


def col(t, c):
    return f"{t}[{c}]"


class Book:
    def __init__(self, path, solved, compat=False):
        self.wb = xlsxwriter.Workbook(path, {"use_future_functions": True})
        self.solved = solved
        self.compat = compat
        self._fmts = {}
        self.ws = {name: self.wb.add_worksheet(name) for name in SHEETS}
        self.key_ws = self.ws["_key"]
        self.key_row = 0
        self.n_checks = Counter()

    # --- formatos
    def f(self, **kw):
        kw = {k: v for k, v in kw.items() if v is not None}
        kw.setdefault("font_size", 10)
        k = tuple(sorted(kw.items()))
        if k not in self._fmts:
            self._fmts[k] = self.wb.add_format(dict(kw))
        return self._fmts[k]

    # --- XLOOKUP (o INDEX/MATCH en modo compatibilidad para probar en LibreOffice)
    def XL(self, val, look, ret, nf=None, mode=0, search=1):
        if not self.compat:
            args = [val, look, ret]
            tail = [nf if nf is not None else "", str(mode) if (mode or search != 1) else "",
                    str(search) if search != 1 else ""]
            while tail and tail[-1] == "":
                tail.pop()
            return f"XLOOKUP({','.join(args + tail)})"
        if search == -1:
            core = f"LOOKUP(2,1/({look}={val}),{ret})"
        elif mode == -1:
            core = f"INDEX({ret},MATCH({val},{look},1))"
        else:
            core = f"INDEX({ret},MATCH({val},{look},0))"
        return f"IFERROR({core},{nf})" if nf is not None else core

    # --- escritura
    def formula(self, ws, r, c, formula, fmt=None):
        if not formula.startswith("="):
            formula = "=" + formula
        if DYNAMIC.search(formula):
            ws.write_dynamic_array_formula(r, c, r, c, formula, fmt)
        else:
            ws.write_formula(r, c, formula, fmt)

    def answer(self, ws, r, c, formula, nf=None, expected=NOCHECK, chk=None, tol=0.005,
               cmp=None, bold=False):
        """Celda de respuesta: amarilla y vacía en la práctica, con fórmula en la solución."""
        if self.solved:
            self.formula(ws, r, c, formula, self.f(num_format=nf, bg_color=C_SOL, border=1,
                                                   border_color="#A9D08E", bold=bold or None))
        else:
            ws.write_blank(r, c, None, self.f(num_format=nf, bg_color=C_INPUT, border=1,
                                              border_color="#BF9000", bold=bold or None))
        if expected is not NOCHECK and chk is not None:
            self.check(ws, chk[0], chk[1], xl_rowcol_to_cell(r, c), expected, tol, cmp)

    def key(self, value):
        if value is None:
            value = ""
        if isinstance(value, dt.date):
            value = serial(value)
        if isinstance(value, str):
            if value:
                self.key_ws.write_string(self.key_row, 0, value)
        else:
            self.key_ws.write_number(self.key_row, 0, value)
        self.key_row += 1
        return f"'_key'!$A${self.key_row}"

    def check(self, ws, r, c, ref, expected, tol=0.005, cmp=None):
        k = self.key(expected)
        if cmp is not None:
            test = cmp.format(ref=ref, key=k)
        elif expected is None or isinstance(expected, str):
            test = f"EXACT({ref},{k})"
        else:
            test = f"ABS({ref}-{k})<={tol}"
        ws.write_formula(r, c, f'=IF(ISBLANK({ref}),"",IF(IFERROR({test},FALSE),"✓","✗"))',
                         self.f(align="center", font_color="#808080"))
        self.n_checks[ws.name] += 1

    def check_cf(self, ws, rng_str):
        ws.conditional_format(rng_str, {"type": "cell", "criteria": "==", "value": '"✓"',
                                        "format": self.f(font_color="#006100", bg_color="#C6EFCE")})
        ws.conditional_format(rng_str, {"type": "cell", "criteria": "==", "value": '"✗"',
                                        "format": self.f(font_color="#9C0006", bg_color="#FFC7CE")})

    def title(self, ws, text, sub=None, ctx=None, width=12):
        ws.set_row(0, 26)
        ws.merge_range(0, 0, 0, width, text, self.f(bold=True, font_size=16, font_color="white",
                                                     bg_color=C_TITLE, valign="vcenter"))
        if sub:
            ws.set_row(1, 45)
            ws.merge_range(1, 0, 1, width, sub, self.f(text_wrap=True, valign="top"))
        if ctx:
            ws.set_row(2, 45)
            ws.merge_range(2, 0, 2, width, ctx, self.f(text_wrap=True, valign="top", italic=True,
                                                        font_color="#595959"))

    def header(self, ws, r, c, labels, hints=None):
        hf = self.f(bold=True, font_color="white", bg_color=C_HDR, text_wrap=True,
                    valign="vcenter", border=1)
        tf = self.f(italic=True, font_color="#595959", bg_color=C_HINT, text_wrap=True,
                    valign="top", border=1, font_size=8)
        for i, lab in enumerate(labels):
            ws.write(r, c + i, lab, hf)
            if hints is not None:
                ws.write(r + 1, c + i, hints[i] if i < len(hints) else "", tf)

    def given(self, ws, r, c, value, nf=None, bold=False):
        fm = self.f(num_format=nf, border=1, border_color="#BFBFBF", bold=bold or None)
        if isinstance(value, dt.date):
            ws.write_datetime(r, c, dt.datetime.combine(value, dt.time()), self.f(
                num_format=nf or NF_DATE, border=1, border_color="#BFBFBF", bold=bold or None))
        else:
            ws.write(r, c, value, fm)

    def note(self, ws, r, c, text, height=None, merge_to=None, rows=1, **kw):
        wrap = merge_to is not None or height is not None or rows > 1
        fm = self.f(text_wrap=wrap or None, valign="top", **kw)
        if merge_to is not None or rows > 1:
            ws.merge_range(r, c, r + rows - 1, c if merge_to is None else merge_to, text, fm)
        else:
            ws.write(r, c, text, fm)
        if height:
            ws.set_row(r, height)

    def close(self):
        self.key_ws.very_hidden()
        self.wb.close()


# ---------------------------------------------------------------------------
# Hojas de datos
# ---------------------------------------------------------------------------

TABLE_DEFS = [
    # hoja, tabla, atributo, columnas (nombre, formato)
    ("Accounts", T_ACC, "accounts", [
        ("AccountID", None), ("AccountName", None), ("Segment", None), ("Region", None),
        ("Industry", None), ("Product", None), ("CSM_ID", None), ("ARR_USD", NF_USD),
        ("LicensedSeats", NF_INT), ("ContractStart", NF_DATE), ("RenewalDate", NF_DATE),
        ("Status", None), ("ChurnDate", NF_DATE), ("ContractualReporting", None)]),
    ("CSMs", T_CSM, "csms", [("CSM_ID", None), ("CSM_Name", None), ("Email", None),
                              ("Team", None), ("Manager", None)]),
    ("Usage_Monthly", T_USE, "usage", [
        ("AccountID", None), ("Month", NF_DATE), ("ActiveUsers", NF_INT), ("Logins", NF_INT),
        ("ReportsRun", NF_INT), ("FeaturesAdopted", None), ("AvgSessionMin", NF_DEC)]),
    ("Support_Tickets", T_TCK, "tickets", [
        ("TicketID", None), ("AccountID", None), ("OpenedDate", NF_DATE), ("ClosedDate", NF_DATE),
        ("Priority", None), ("Category", None), ("CSAT", None)]),
    ("NPS", T_NPS, "nps", [("SurveyID", None), ("AccountID", None), ("ResponseDate", NF_DATE),
                           ("Score", None)]),
    ("Report_Catalog", T_REP, "reports", [
        ("ReportID", None), ("ReportName", None), ("Type", None), ("Audience", None),
        ("Frequency", None), ("Source", None), ("Owner", None), ("CreatedDate", NF_DATE),
        ("LastModified", NF_DATE), ("Recipients", None), ("MaintHrsPerMonth", NF_DEC)]),
    ("Report_Usage_Log", T_LOG, "log", [
        ("LogID", None), ("Date", NF_DATE), ("ReportID", None), ("UserEmail", None),
        ("UserRole", None), ("Action", None), ("DurationSec", NF_INT)]),
]


def write_data_sheets(bk, D):
    for sheet, tname, attr, cols in TABLE_DEFS:
        ws = bk.ws[sheet]
        rows = getattr(D, attr)
        data = []
        for r in rows:
            line = []
            for cname, _ in cols:
                v = r[cname]
                if isinstance(v, dt.date):
                    v = dt.datetime.combine(v, dt.time())
                line.append(v)
            data.append(line)
        ws.add_table(0, 0, len(data), len(cols) - 1, {
            "name": tname, "style": "Table Style Medium 2", "data": data,
            "columns": [{"header": c, "format": bk.f(num_format=nf)} for c, nf in cols]})
        for i, (c, _) in enumerate(cols):
            width = max(len(c) + 3, min(38, max((len(str(r[c])) for r in rows), default=8) + 2))
            ws.set_column(i, i, width)
        ws.freeze_panes(1, 0)
        ws.set_tab_color("#808080")


def write_csvs(D, outdir):
    os.makedirs(outdir, exist_ok=True)
    for sheet, _, attr, cols in TABLE_DEFS:
        with open(os.path.join(outdir, f"{sheet}.csv"), "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow([c for c, _ in cols])
            for r in getattr(D, attr):
                w.writerow(["" if r[c] is None else (r[c].isoformat() if isinstance(r[c], dt.date)
                                                     else r[c]) for c, _ in cols])


# ---------------------------------------------------------------------------
# LEEME y diccionario
# ---------------------------------------------------------------------------

FUNC_ES = [
    ("XLOOKUP", "BUSCARX"), ("VLOOKUP", "BUSCARV"), ("HLOOKUP", "BUSCARH"), ("INDEX", "INDICE"),
    ("MATCH", "COINCIDIR"), ("XMATCH", "COINCIDIRX"), ("SUMIFS", "SUMAR.SI.CONJUNTO"),
    ("COUNTIFS", "CONTAR.SI.CONJUNTO"), ("AVERAGEIFS", "PROMEDIO.SI.CONJUNTO"),
    ("MAXIFS / MINIFS", "MAX.SI.CONJUNTO / MIN.SI.CONJUNTO"), ("SUMPRODUCT", "SUMAPRODUCTO"),
    ("IFS", "SI.CONJUNTO"), ("SWITCH", "CAMBIAR"), ("IFERROR", "SI.ERROR"), ("TRIM", "ESPACIOS"),
    ("PROPER / UPPER / LOWER", "NOMPROPIO / MAYUSC / MINUSC"), ("SUBSTITUTE", "SUSTITUIR"),
    ("NUMBERVALUE", "VALOR.NUMERO"), ("LEFT / MID / RIGHT", "IZQUIERDA / EXTRAE / DERECHA"),
    ("EOMONTH / EDATE", "FIN.MES / FECHA.MES"), ("RANK.EQ", "JERARQUIA.EQV"),
    ("FILTER / UNIQUE / SORT / SORTBY", "FILTRAR / UNICOS / ORDENAR / ORDENARPOR"),
    ("TEXTJOIN", "UNIRCADENAS"), ("LET", "LET"), ("TAKE / HSTACK", "TOMAR / APILARH"),
    ("SEQUENCE", "SECUENCIA"), ("GETPIVOTDATA", "IMPORTARDATOSDINAMICOS"),
]


def sheet_leeme(bk, D):
    ws = bk.ws["LEEME"]
    ws.set_column(0, 0, 3)
    ws.set_column(1, 1, 30)
    ws.set_column(2, 2, 16)
    ws.set_column(3, 3, 62)
    ws.set_column(4, 6, 12)
    modo = "SOLUCIONES (todas las fórmulas resueltas)" if bk.solved else "PRÁCTICA (llena las celdas amarillas)"
    ws.set_row(0, 30)
    ws.merge_range(0, 0, 0, 6, "Práctica de Excel · Analyst, Customer Success Operations",
                   bk.f(bold=True, font_size=18, font_color="white", bg_color=C_TITLE, valign="vcenter"))
    ws.merge_range(1, 0, 1, 6, f"Versión: {modo}", bk.f(bold=True, font_color=C_TITLE))
    bk.note(ws, 2, 1,
            "Escenario: eres analista de CS Ops en una empresa SaaS B2B. Liderazgo quiere (1) entender la salud "
            "de las cuentas para que los CSMs prioricen, (2) auditar los reportes actuales para decidir qué "
            "mantener, consolidar, automatizar o retirar, y (3) tener datos limpios y confiables para QBRs/EBRs, "
            "renovaciones y reportes contractuales. Todos los datos son ficticios.",
            merge_to=6, height=48)

    bk.note(ws, 4, 1, "Parámetros (usados en todas las fórmulas)", bold=True, font_color=C_TITLE)
    pf = bk.f(bold=True, bg_color="#DDEBF7", border=1)
    ws.write(5, 1, "AsOfDate (fecha de corte)")
    ws.write_datetime(5, 2, dt.datetime.combine(D.as_of, dt.time()), bk.f(num_format=NF_DATE, bold=True,
                                                                          bg_color="#DDEBF7", border=1))
    ws.write(5, 3, "Nombre definido AsOfDate. Se usa en lugar de TODAY() para que los resultados sean reproducibles.")
    ws.write(6, 1, "LastMonth (último mes cerrado)")
    ws.write_formula(6, 2, "=DATE(YEAR(AsOfDate),MONTH(AsOfDate),1)",
                     bk.f(num_format=NF_DATE, bold=True, bg_color="#DDEBF7", border=1),
                     serial(D.last_month))
    ws.write(6, 3, "Nombre definido LastMonth = primer día del mes de AsOfDate.")
    ws.write(7, 1, "CostPerHour (USD)")
    ws.write_number(7, 2, 35, pf)
    ws.write(7, 3, "Costo por hora de analista; se usa para estimar el costo de mantener un reporte.")
    bk.wb.define_name("AsOfDate", "=LEEME!$C$6")
    bk.wb.define_name("LastMonth", "=LEEME!$C$7")
    bk.wb.define_name("CostPerHour", "=LEEME!$C$8")

    r = 9
    bk.note(ws, r, 1, "Cómo usar este libro", bold=True, font_color=C_TITLE)
    steps = [
        "1. Cada hoja E# es un ejercicio. Lee el objetivo (fila 2), el contexto de negocio (fila 3) y la pista "
        "gris debajo de cada encabezado.",
        "2. Escribe tus fórmulas en las celdas AMARILLAS. Las columnas ✓/✗ se validan solas contra la respuesta "
        "correcta (oculta). Copia/arrastra la fórmula de la primera fila hacia abajo.",
        "3. Si te atoras, abre CS_Ops_Soluciones.xlsx (mismo set de datos) y compara tu fórmula.",
        "4. Los datos están en las hojas grises como Tablas de Excel (tbl_Accounts, tbl_Usage, etc.). "
        "Úsalas con referencias estructuradas: tbl_Accounts[ARR_USD]. Revisa la hoja Diccionario.",
        "5. Las verificaciones asumen los parámetros y pesos por defecto. Si los cambias para experimentar, "
        "algunas ✓ pasarán a ✗ (eso es normal).",
        "6. Regenera el archivo con otra semilla (python generar_practica_excel.py --semilla 123) para practicar "
        "de nuevo con números diferentes.",
        "7. Excel en español: los nombres de funciones se traducen solos al abrir el archivo y el separador "
        "puede ser ';' en vez de ','. Abajo a la derecha tienes la tabla de traducción.",
    ]
    for i, s in enumerate(steps):
        bk.note(ws, r + 1 + i, 1, s, merge_to=6, height=30)

    r = r + len(steps) + 2
    bk.note(ws, r, 1, "Ejercicios y progreso", bold=True, font_color=C_TITLE)
    bk.header(ws, r + 1, 1, ["Hoja", "Habilidad clave", "Qué practicas / para qué sirve en el puesto",
                             "✓", "Total", "Avance"])
    desc = {
        "E1_Lookups": ("Lookups", "VLOOKUP, INDEX/MATCH, XLOOKUP anidado, búsqueda aproximada (tiers), "
                                  "búsqueda inversa (último registro), IFERROR y TRIM."),
        "E2_Agregaciones": ("SUMIFS/COUNTIFS", "Matrices ARR segmento×región, book of business por CSM, "
                                               "tickets por prioridad, NPS. Base de cualquier reporte."),
        "E3_Health_Score": ("Health score", "Construir un health score ponderado (adopción, tendencia, "
                                            "soporte, NPS), categorías Green/Yellow/Red y riesgo de renovación."),
        "E4_Auditoria_Reportes": ("Auditoría de reportes", "Uso real de cada reporte, usuarios únicos, "
                                                           "última consulta, costo por vista y recomendación "
                                                           "(retirar / consolidar / automatizar / mantener)."),
        "E5_Limpieza_Datos": ("Calidad de datos", "Limpiar un export del CRM: espacios, mayúsculas, montos y "
                                                  "fechas en texto, duplicados, IDs inexistentes y "
                                                  "diferencias vs. el maestro."),
        "E6_Pivots": ("Tablas dinámicas", "Construir pivots (agrupación por mes, % del total, campos "
                                          "calculados, slicers) y validar con GETPIVOTDATA."),
        "E7_Dinamicas_365": ("Matrices dinámicas", "FILTER, UNIQUE, SORT, XLOOKUP de varias columnas, "
                                                   "TEXTJOIN, LET, SEQUENCE."),
        "E8_Tendencias": ("Tendencias y churn", "Uso mensual, MoM %, promedio móvil, churn y GRR, pipeline de "
                                                "renovaciones con ARR en riesgo."),
        "E9_Dashboard_CSM": ("Dashboard", "Tablero para un CSM con lista desplegable, KPIs, top 5 cuentas en "
                                          "riesgo y gráfica."),
    }
    rr = r + 2
    first = rr
    for name in EXERCISES:
        skill, what = desc[name]
        ws.write_url(rr, 1, f"internal:'{name}'!A1", bk.f(font_color="blue", underline=1), string=name)
        ws.write(rr, 2, skill, bk.f(bold=True, text_wrap=True))
        ws.write(rr, 3, what, bk.f(text_wrap=True))
        ws.write_formula(rr, 4, f"=COUNTIF('{name}'!$A$1:$BZ$400,\"✓\")", bk.f(align="center"))
        ws.write_number(rr, 5, 0, bk.f(align="center"))  # se rellena al final
        ws.write_formula(rr, 6, f"=IFERROR(E{rr + 1}/F{rr + 1},0)", bk.f(num_format="0%"))
        ws.set_row(rr, 30)
        rr += 1
    ws.write(rr, 3, "TOTAL", bk.f(bold=True, align="right"))
    ws.write_formula(rr, 4, f"=SUM(E{first + 1}:E{rr})", bk.f(bold=True, align="center"))
    ws.write_formula(rr, 5, f"=SUM(F{first + 1}:F{rr})", bk.f(bold=True, align="center"))
    ws.write_formula(rr, 6, f"=IFERROR(E{rr + 1}/F{rr + 1},0)", bk.f(bold=True, num_format="0%"))
    ws.conditional_format(first, 6, rr, 6, {"type": "data_bar", "bar_color": "#63BE7B",
                                            "bar_solid": True, "min_type": "num", "min_value": 0,
                                            "max_type": "num", "max_value": 1})
    bk._progress_rows = (first, EXERCISES)

    rr += 2
    bk.note(ws, rr, 1, "Extras (sin verificación automática)", bold=True, font_color=C_TITLE)
    extras = [("E10_PowerQuery", "Importar y combinar los CSV con Power Query; refrescar."),
              ("Entrevista", "Preguntas técnicas y de negocio para practicar en voz alta (en inglés)."),
              ("Diccionario", "Qué contiene cada tabla y columna.")]
    for i, (n, d) in enumerate(extras):
        ws.write_url(rr + 1 + i, 1, f"internal:'{n}'!A1", bk.f(font_color="blue", underline=1), string=n)
        ws.write(rr + 1 + i, 3, d)

    rr += 6
    bk.note(ws, rr, 1, "Atajos útiles", bold=True, font_color=C_TITLE)
    shortcuts = [("Ctrl + T", "Convertir rango en Tabla"), ("Alt, N, V", "Insertar tabla dinámica"),
                 ("F4", "Fijar referencia ($A$1)"), ("Ctrl + Shift + L", "Activar filtros"),
                 ("Ctrl + Flecha", "Ir al final de los datos"), ("Ctrl + Shift + Flecha", "Seleccionar hasta el final"),
                 ("Alt + =", "Autosuma"), ("Ctrl + `", "Mostrar fórmulas"),
                 ("F9 (dentro de una fórmula)", "Evaluar una parte de la fórmula"),
                 ("Fórmulas > Evaluar fórmula", "Depurar fórmulas paso a paso"),
                 ("Ctrl + Alt + F5", "Actualizar todo (pivots / Power Query)")]
    for i, (k, d) in enumerate(shortcuts):
        ws.write(rr + 1 + i, 1, k, bk.f(bold=True))
        ws.write(rr + 1 + i, 3, d)

    # traducción de funciones
    rr += len(shortcuts) + 2
    bk.note(ws, rr, 1, "Funciones: inglés → español (Excel las traduce solo)", bold=True, font_color=C_TITLE)
    bk.header(ws, rr + 1, 1, ["Inglés", "", "Español"])
    for i, (en, es) in enumerate(FUNC_ES):
        ws.write(rr + 2 + i, 1, en, bk.f(border=1, border_color="#BFBFBF"))
        ws.write(rr + 2 + i, 3, es, bk.f(border=1, border_color="#BFBFBF"))
    ws.activate()
    ws.set_tab_color(C_TITLE)


def finalize_progress(bk):
    ws = bk.ws["LEEME"]
    first, names = bk._progress_rows
    for i, name in enumerate(names):
        ws.write_number(first + i, 5, bk.n_checks[name], bk.f(align="center"))


DICT = [
    ("tbl_Accounts", "Accounts", "Maestro de cuentas (1 fila por cliente)", [
        ("AccountID", "ID único de la cuenta (llave primaria)"), ("AccountName", "Nombre del cliente"),
        ("Segment", "Enterprise / Mid-Market / SMB"), ("Region", "Región comercial"),
        ("Industry", "Industria"), ("Product", "Producto principal contratado"),
        ("CSM_ID", "CSM asignado (llave hacia tbl_CSMs)"), ("ARR_USD", "Ingreso recurrente anual"),
        ("LicensedSeats", "Licencias contratadas"), ("ContractStart", "Inicio de contrato"),
        ("RenewalDate", "Próxima renovación (o fecha de churn si Status = Churned)"),
        ("Status", "Active / Churned"), ("ChurnDate", "Fecha en que se perdió la cuenta"),
        ("ContractualReporting", "Si el contrato obliga a enviar reportes de uso (Monthly/Quarterly/None)")]),
    ("tbl_CSMs", "CSMs", "Customer Success Managers", [
        ("CSM_ID", "ID del CSM"), ("CSM_Name", "Nombre"), ("Email", "Correo"), ("Team", "Segmento que atiende"),
        ("Manager", "Líder")]),
    ("tbl_Usage", "Usage_Monthly", "Uso del producto por cuenta y mes (12 meses)", [
        ("AccountID", "Cuenta"), ("Month", "Primer día del mes"), ("ActiveUsers", "Usuarios activos en el mes"),
        ("Logins", "Inicios de sesión"), ("ReportsRun", "Reportes ejecutados dentro del producto"),
        ("FeaturesAdopted", "Funcionalidades clave usadas (1-10)"), ("AvgSessionMin", "Duración media de sesión")]),
    ("tbl_Tickets", "Support_Tickets", "Tickets de soporte (último año)", [
        ("TicketID", "ID"), ("AccountID", "Cuenta"), ("OpenedDate", "Apertura"),
        ("ClosedDate", "Cierre (vacío = abierto)"), ("Priority", "P1 (crítico) a P4"), ("Category", "Tipo"),
        ("CSAT", "Satisfacción 1-5 (vacío = sin respuesta)")]),
    ("tbl_NPS", "NPS", "Respuestas de encuesta NPS (algunas > 12 meses)", [
        ("SurveyID", "ID"), ("AccountID", "Cuenta"), ("ResponseDate", "Fecha"),
        ("Score", "0-10. Promotor 9-10, Pasivo 7-8, Detractor 0-6")]),
    ("tbl_Reports", "Report_Catalog", "Inventario de reportes que mantiene CS Ops", [
        ("ReportID", "ID"), ("ReportName", "Nombre"), ("Type", "Standard / Custom / Contractual"),
        ("Audience", "Para quién es"), ("Frequency", "Cada cuánto se actualiza/envía"),
        ("Source", "Sistema origen"), ("Owner", "Analista dueño"), ("CreatedDate", "Creación"),
        ("LastModified", "Última modificación"), ("Recipients", "Personas en la lista de distribución"),
        ("MaintHrsPerMonth", "Horas al mes para mantenerlo")]),
    ("tbl_ReportLog", "Report_Usage_Log", "Log de consultas de reportes (últimos 180 días)", [
        ("LogID", "ID"), ("Date", "Fecha"), ("ReportID", "Reporte"), ("UserEmail", "Usuario"),
        ("UserRole", "Rol del usuario"), ("Action", "View / Export / Subscribe"),
        ("DurationSec", "Segundos en el reporte")]),
]


def sheet_dictionary(bk, D):
    ws = bk.ws["Diccionario"]
    bk.title(ws, "Diccionario de datos", "Antes de analizar, entiende la granularidad de cada tabla "
             "(¿una fila = qué?) y cuáles son las llaves para relacionarlas.", width=4)
    ws.set_column(0, 0, 18)
    ws.set_column(1, 1, 20)
    ws.set_column(2, 2, 24)
    ws.set_column(3, 3, 70)
    ws.set_column(4, 4, 10)
    bk.header(ws, 4, 0, ["Tabla", "Hoja", "Columna", "Descripción", "Filas"])
    r = 5
    counts = {"tbl_Accounts": len(D.accounts), "tbl_CSMs": len(D.csms), "tbl_Usage": len(D.usage),
              "tbl_Tickets": len(D.tickets), "tbl_NPS": len(D.nps), "tbl_Reports": len(D.reports),
              "tbl_ReportLog": len(D.log)}
    for t, sheet, desc, cols in DICT:
        ws.write(r, 0, t, bk.f(bold=True, bg_color="#DDEBF7"))
        ws.write_url(r, 1, f"internal:'{sheet}'!A1", bk.f(font_color="blue", underline=1,
                                                         bg_color="#DDEBF7"), string=sheet)
        ws.write(r, 2, "", bk.f(bg_color="#DDEBF7"))
        ws.write(r, 3, desc, bk.f(italic=True, bg_color="#DDEBF7"))
        ws.write(r, 4, counts[t], bk.f(bg_color="#DDEBF7", num_format=NF_INT))
        r += 1
        for c, d in cols:
            ws.write(r, 2, c)
            ws.write(r, 3, d)
            r += 1
    r += 1
    bk.note(ws, r, 0, "Relaciones: tbl_Accounts[AccountID] 1→N tbl_Usage / tbl_Tickets / tbl_NPS · "
                      "tbl_CSMs[CSM_ID] 1→N tbl_Accounts · tbl_Reports[ReportID] 1→N tbl_ReportLog",
            merge_to=4, bold=True, height=30)


# ---------------------------------------------------------------------------
# E1 - Lookups
# ---------------------------------------------------------------------------

def tier_of(arr):
    t = None
    for lo, name in TIERS:
        if arr >= lo:
            t = name
    return t


def sheet_e1(bk, D):
    ws = bk.ws["E1_Lookups"]
    rng = random.Random(D.seed + 1)
    bk.title(ws, "E1 · Lookups: traer datos de otras tablas",
             "Objetivo: un CSM te manda una lista de IDs (copiada de un correo, con espacios y minúsculas) y "
             "necesita nombre, segmento, CSM, ARR, tier, uso del último mes y días para la renovación. "
             "Practica las 3 familias de búsqueda y cuándo usar cada una.",
             "Contexto: es el trabajo diario de CS Ops — enriquecer listas con datos del CRM. XLOOKUP es el "
             "estándar hoy, pero en entrevistas y archivos heredados verás VLOOKUP e INDEX/MATCH.", width=18)
    ids = [a["AccountID"] for a in rng.sample(D.accounts, 12)]
    ids.insert(3, f"{rng.choice(D.accounts)['AccountID'].lower()} ")
    ids.insert(7, "ACC-9999")
    ids.insert(11, f" {rng.choice(D.accounts)['AccountID']}")
    labels = ["AccountID (lista del CSM)", "AccountName", "Segment", "CSM_Name", "ARR_USD", "Tier (aprox.)",
              "Tier (XLOOKUP)", "ActiveUsers último registro", "Días a renovación"]
    hints = ["Dato dado (ojo: espacios y minúsculas)",
             "VLOOKUP(TRIM(A7), tbl_Accounts, 2, FALSE) + IFERROR → \"No encontrado\"",
             "INDEX(tbl_Accounts[Segment], MATCH(TRIM(A7), tbl_Accounts[AccountID], 0)) + IFERROR → \"No encontrado\"",
             "XLOOKUP anidado: primero el CSM_ID en tbl_Accounts y luego el nombre en tbl_CSMs. IFERROR → \"No encontrado\"",
             "XLOOKUP con if_not_found = 0",
             "Si ARR=0 → \"\"; si no VLOOKUP(E7, $T$7:$U$10, 2, TRUE) (búsqueda aproximada)",
             "Igual pero con XLOOKUP(E7, $T$7:$T$10, $U$7:$U$10, , -1)",
             "XLOOKUP con search_mode -1 (de abajo hacia arriba) en tbl_Usage → último mes con datos. "
             "if_not_found → \"No encontrado\"",
             "IFERROR(XLOOKUP(...RenewalDate) - AsOfDate, \"\")"]
    bk.header(ws, 4, 0, labels, hints)
    chk_c = 10
    bk.header(ws, 4, chk_c, ["✓B", "✓C", "✓D", "✓E", "✓F", "✓G", "✓H", "✓I"])
    ws.set_row(5, 62)
    ws.set_column(0, 0, 16)
    ws.set_column(1, 8, 17)
    ws.set_column(9, 9, 2)
    ws.set_column(chk_c, chk_c + 7, 4)
    ws.set_column(18, 18, 2)
    ws.set_column(19, 20, 12)

    last_usage = {}
    for u in D.usage:
        last_usage[u["AccountID"]] = u["ActiveUsers"]
    csm_name = {c["CSM_ID"]: c["CSM_Name"] for c in D.csms}
    A = T_ACC
    for i, raw in enumerate(ids):
        r = 6 + i
        x = f"A{r + 1}"
        bk.given(ws, r, 0, raw)
        acc = D.acc_by_id.get(raw.strip().upper())
        nf_txt = '"No encontrado"'
        bk.answer(ws, r, 1, f"IFERROR(VLOOKUP(TRIM({x}),{A},2,FALSE),{nf_txt})",
                  expected=acc["AccountName"] if acc else "No encontrado", chk=(r, chk_c))
        bk.answer(ws, r, 2, f"IFERROR(INDEX({col(A, 'Segment')},MATCH(TRIM({x}),{col(A, 'AccountID')},0)),{nf_txt})",
                  expected=acc["Segment"] if acc else "No encontrado", chk=(r, chk_c + 1))
        inner = bk.XL(f"TRIM({x})", col(A, "AccountID"), col(A, "CSM_ID"))
        bk.answer(ws, r, 3, f"IFERROR({bk.XL(inner, col(T_CSM, 'CSM_ID'), col(T_CSM, 'CSM_Name'))},{nf_txt})",
                  expected=csm_name[acc["CSM_ID"]] if acc else "No encontrado", chk=(r, chk_c + 2))
        bk.answer(ws, r, 4, bk.XL(f"TRIM({x})", col(A, "AccountID"), col(A, "ARR_USD"), nf="0"), nf=NF_USD,
                  expected=acc["ARR_USD"] if acc else 0, chk=(r, chk_c + 3))
        tier = tier_of(acc["ARR_USD"]) if acc else ""
        bk.answer(ws, r, 5, f'IF(E{r + 1}=0,"",VLOOKUP(E{r + 1},$T$7:$U$10,2,TRUE))',
                  expected=tier, chk=(r, chk_c + 4))
        bk.answer(ws, r, 6, f'IF(E{r + 1}=0,"",{bk.XL(f"E{r + 1}", "$T$7:$T$10", "$U$7:$U$10", mode=-1)})',
                  expected=tier, chk=(r, chk_c + 5))
        bk.answer(ws, r, 7, bk.XL(f"TRIM({x})", col(T_USE, "AccountID"), col(T_USE, "ActiveUsers"),
                                  nf=nf_txt, search=-1), nf=NF_INT,
                  expected=last_usage[acc["AccountID"]] if acc else "No encontrado", chk=(r, chk_c + 6))
        bk.answer(ws, r, 8, f'IFERROR({bk.XL(f"TRIM({x})", col(A, "AccountID"), col(A, "RenewalDate"))}-AsOfDate,"")',
                  nf=NF_INT, expected=(acc["RenewalDate"] - D.as_of).days if acc else "", chk=(r, chk_c + 7))
    last = 6 + len(ids)
    bk.check_cf(ws, f"K7:R{last}")

    bk.header(ws, 4, 19, ["Min ARR", "Tier"])
    ws.merge_range(5, 19, 5, 20, "Tabla de tiers: ordenada ascendente (requisito de la búsqueda aproximada)",
                   bk.f(italic=True, font_size=8, text_wrap=True, bg_color=C_HINT))
    for i, (lo, name) in enumerate(TIERS):
        bk.given(ws, 6 + i, 19, lo, NF_USD)
        bk.given(ws, 6 + i, 20, name)

    r = last + 1
    bk.note(ws, r, 0, "Para pensar (entrevista): ¿por qué VLOOKUP se rompe si alguien inserta una columna en "
                      "tbl_Accounts? ¿Qué ventajas tiene XLOOKUP (búsqueda a la izquierda, if_not_found, búsqueda "
                      "inversa, coincidencia aproximada sin ordenar)? ¿Qué pasa con el ID que tiene espacios si no "
                      "usas TRIM?", merge_to=12, height=45, italic=True)
    ws.freeze_panes(6, 1)


# ---------------------------------------------------------------------------
# E2 - Agregaciones
# ---------------------------------------------------------------------------

def sheet_e2(bk, D):
    ws = bk.ws["E2_Agregaciones"]
    bk.title(ws, "E2 · Agregaciones condicionales (SUMIFS, COUNTIFS, AVERAGEIFS, MAXIFS, SUMPRODUCT)",
             "Objetivo: responder con fórmulas las preguntas típicas de liderazgo: ¿cuánto ARR activo tenemos por "
             "segmento y región? ¿cómo está el book of business de cada CSM? ¿cómo va el backlog de soporte? "
             "¿cuál es nuestro NPS?",
             "Tip: estas mismas preguntas se pueden contestar con una tabla dinámica (E6). Saber hacerlas con "
             "fórmulas te permite construir reportes que se actualizan solos y validar los pivots.", width=16)
    ws.set_column(0, 0, 22)
    ws.set_column(1, 7, 15)
    ws.set_column(8, 8, 2)
    ws.set_column(9, 16, 4.5)
    A = T_ACC
    active = [a for a in D.accounts if a["Status"] == "Active"]

    # Bloque 1: matriz ARR
    bk.note(ws, 3, 0, "1) ARR activo por Segmento × Región", bold=True, font_color=C_TITLE)
    bk.header(ws, 4, 0, ["Segment \\ Region"] + REGIONS + ["Total"],
              ["", "SUMIFS(tbl_Accounts[ARR_USD], ...[Segment], $A7, ...[Region], B$5, ...[Status], \"Active\")",
               "Fija fila/columna con $ para arrastrar", "", "", "SUM de la fila"])
    ws.set_row(5, 50)
    bk.header(ws, 4, 9, ["Chk"] * 5)
    for i, seg in enumerate(SEGMENTS + ["Total"]):
        r = 6 + i
        bk.given(ws, r, 0, seg, bold=True)
        for j, reg in enumerate(REGIONS + ["Total"]):
            c = 1 + j
            if seg != "Total" and reg != "Total":
                fml = (f'SUMIFS({col(A, "ARR_USD")},{col(A, "Segment")},$A{r + 1},{col(A, "Region")},'
                       f'{xl_col_to_name(c)}$5,{col(A, "Status")},"Active")')
                exp = sum(a["ARR_USD"] for a in active if a["Segment"] == seg and a["Region"] == reg)
            elif reg == "Total" and seg != "Total":
                fml = f"SUM(B{r + 1}:E{r + 1})"
                exp = sum(a["ARR_USD"] for a in active if a["Segment"] == seg)
            elif seg == "Total" and reg != "Total":
                fml = f"SUM({xl_col_to_name(c)}7:{xl_col_to_name(c)}9)"
                exp = sum(a["ARR_USD"] for a in active if a["Region"] == reg)
            else:
                fml = "SUM(F7:F9)"
                exp = sum(a["ARR_USD"] for a in active)
            bk.answer(ws, r, c, fml, nf=NF_USD, expected=exp, chk=(r, 9 + j), bold=seg == "Total" or reg == "Total")
    bk.check_cf(ws, "J7:N10")

    # Bloque 2: book of business por CSM
    r0 = 12
    bk.note(ws, r0, 0, "2) Book of business por CSM (solo cuentas activas)", bold=True, font_color=C_TITLE)
    bk.header(ws, r0 + 1, 0, ["CSM_ID", "CSM_Name", "# Cuentas", "ARR", "ARR promedio", "Cuenta más grande (ARR)",
                              "Próxima renovación", "% del ARR total"],
              ["Dado", "XLOOKUP en tbl_CSMs", "COUNTIFS con 2 criterios", "SUMIFS", "AVERAGEIFS (IFERROR → 0)",
               "MAXIFS", "MINIFS sobre RenewalDate", "D / SUM($D$...) "])
    ws.set_row(r0 + 2, 40)
    bk.header(ws, r0 + 1, 9, ["✓B", "✓C", "✓D", "✓E", "✓F", "✓G", "✓H"])
    total_arr = sum(a["ARR_USD"] for a in active)
    rows = list(range(r0 + 3, r0 + 3 + len(D.csms)))
    for c_, r in zip(D.csms, rows):
        x = f"$A{r + 1}"
        mine = [a for a in active if a["CSM_ID"] == c_["CSM_ID"]]
        bk.given(ws, r, 0, c_["CSM_ID"])
        bk.answer(ws, r, 1, bk.XL(x, col(T_CSM, "CSM_ID"), col(T_CSM, "CSM_Name")),
                  expected=c_["CSM_Name"], chk=(r, 9))
        crit = f'{col(A, "CSM_ID")},{x},{col(A, "Status")},"Active"'
        bk.answer(ws, r, 2, f"COUNTIFS({crit})", expected=len(mine), chk=(r, 10))
        bk.answer(ws, r, 3, f'SUMIFS({col(A, "ARR_USD")},{crit})', nf=NF_USD,
                  expected=sum(a["ARR_USD"] for a in mine), chk=(r, 11))
        bk.answer(ws, r, 4, f'IFERROR(AVERAGEIFS({col(A, "ARR_USD")},{crit}),0)', nf=NF_USD,
                  expected=(sum(a["ARR_USD"] for a in mine) / len(mine)) if mine else 0, chk=(r, 12))
        bk.answer(ws, r, 5, f'MAXIFS({col(A, "ARR_USD")},{crit})', nf=NF_USD,
                  expected=max((a["ARR_USD"] for a in mine), default=0), chk=(r, 13))
        bk.answer(ws, r, 6, f'MINIFS({col(A, "RenewalDate")},{crit})', nf=NF_DATE,
                  expected=min((a["RenewalDate"] for a in mine), default=EPOCH), chk=(r, 14))
        bk.answer(ws, r, 7, f"D{r + 1}/SUM($D${rows[0] + 1}:$D${rows[-1] + 1})", nf=NF_PCT,
                  expected=sum(a["ARR_USD"] for a in mine) / total_arr, chk=(r, 15), tol=0.0005)
    bk.check_cf(ws, f"J{rows[0] + 1}:P{rows[-1] + 1}")

    # Bloque 3: tickets por prioridad
    r0 = rows[-1] + 3
    bk.note(ws, r0, 0, "3) Soporte por prioridad (todos los tickets del año)", bold=True, font_color=C_TITLE)
    bk.header(ws, r0 + 1, 0, ["Priority", "Total tickets", "Abiertos", "% abiertos", "Días prom. de resolución",
                              "CSAT promedio"],
              ["Dado", "COUNTIFS", "COUNTIFS con ClosedDate = \"\" (vacío)", "C / B",
               "SUMPRODUCT((Priority=A)*(ClosedDate<>\"\")*(ClosedDate-OpenedDate)) / COUNTIFS(..., ClosedDate, \"<>\")",
               "AVERAGEIFS (ignora vacíos) + IFERROR"])
    ws.set_row(r0 + 2, 62)
    bk.header(ws, r0 + 1, 9, ["✓B", "✓C", "✓D", "✓E", "✓F"])
    T = T_TCK
    prow = list(range(r0 + 3, r0 + 7))
    for p, r in zip(PRIORITIES, prow):
        x = f"$A{r + 1}"
        ts = [t for t in D.tickets if t["Priority"] == p]
        closed = [t for t in ts if t["ClosedDate"]]
        csat = [t["CSAT"] for t in ts if t["CSAT"] is not None]
        bk.given(ws, r, 0, p)
        bk.answer(ws, r, 1, f'COUNTIFS({col(T, "Priority")},{x})', expected=len(ts), chk=(r, 9))
        bk.answer(ws, r, 2, f'COUNTIFS({col(T, "Priority")},{x},{col(T, "ClosedDate")},"")',
                  expected=len(ts) - len(closed), chk=(r, 10))
        bk.answer(ws, r, 3, f"IFERROR(C{r + 1}/B{r + 1},0)", nf=NF_PCT,
                  expected=(len(ts) - len(closed)) / len(ts) if ts else 0, chk=(r, 11), tol=0.0005)
        bk.answer(ws, r, 4, f'IFERROR(SUMPRODUCT(({col(T, "Priority")}={x})*({col(T, "ClosedDate")}<>"")*'
                            f'({col(T, "ClosedDate")}-{col(T, "OpenedDate")}))/'
                            f'COUNTIFS({col(T, "Priority")},{x},{col(T, "ClosedDate")},"<>"),0)', nf=NF_DEC,
                  expected=(sum((t["ClosedDate"] - t["OpenedDate"]).days for t in closed) / len(closed))
                  if closed else 0, chk=(r, 12), tol=0.05)
        bk.answer(ws, r, 5, f'IFERROR(AVERAGEIFS({col(T, "CSAT")},{col(T, "Priority")},{x}),"")', nf=NF_DEC2,
                  expected=(sum(csat) / len(csat)) if csat else "", chk=(r, 13), tol=0.005)
    bk.check_cf(ws, f"J{prow[0] + 1}:N{prow[-1] + 1}")

    # Bloque 4: NPS
    r0 = prow[-1] + 3
    bk.note(ws, r0, 0, "4) NPS global (todas las respuestas)", bold=True, font_color=C_TITLE)
    bk.header(ws, r0 + 1, 0, ["Métrica", "Valor", "Pista"])
    ws.set_column(2, 2, 15)
    scores = [n["Score"] for n in D.nps]
    prom = sum(1 for s in scores if s >= 9)
    pas = sum(1 for s in scores if 7 <= s <= 8)
    det = sum(1 for s in scores if s <= 6)
    S = col(T_NPS, "Score")
    items = [("Promotores (9-10)", f'COUNTIF({S},">=9")', prom, "COUNTIF >=9"),
             ("Pasivos (7-8)", f'COUNTIFS({S},">=7",{S},"<=8")', pas, "COUNTIFS con 2 criterios sobre la misma columna"),
             ("Detractores (0-6)", f'COUNTIF({S},"<=6")', det, "COUNTIF <=6"),
             ("Total respuestas", f"SUM(B{r0 + 3}:B{r0 + 5})", len(scores), "SUM"),
             ("NPS", f"(B{r0 + 3}-B{r0 + 5})/B{r0 + 6}*100", (prom - det) / len(scores) * 100,
              "(%Promotores − %Detractores) × 100. Rango −100 a 100")]
    for i, (lab, fml, exp, hint) in enumerate(items):
        r = r0 + 2 + i
        bk.given(ws, r, 0, lab, bold=lab == "NPS")
        bk.answer(ws, r, 1, fml, nf=NF_DEC if lab == "NPS" else NF_INT, expected=exp, chk=(r, 9), tol=0.05)
        ws.write(r, 2, hint, bk.f(italic=True, font_color="#595959", font_size=8))
    bk.check_cf(ws, f"J{r0 + 3}:J{r0 + 7}")
    ws.freeze_panes(3, 1)


# ---------------------------------------------------------------------------
# E3 - Health Score
# ---------------------------------------------------------------------------

WEIGHTS = [0.35, 0.25, 0.20, 0.20]


def s_adoption(u):
    return 100 if u >= 0.8 else 75 if u >= 0.6 else 50 if u >= 0.4 else 25 if u >= 0.2 else 0


def s_trend(t):
    return 100 if t >= 0.1 else 75 if t >= 0 else 50 if t >= -0.1 else 25 if t >= -0.25 else 0


def s_support(k):
    return 100 if k == 0 else 60 if k == 1 else 30 if k <= 3 else 0


def s_nps(n):
    if n == "":
        return 50
    return 100 if n >= 9 else 70 if n >= 7 else 40 if n >= 5 else 10


def compute_health(D):
    usage = {(u["AccountID"], u["Month"]): u for u in D.usage}
    m = D.months
    cut = edate(D.as_of, -12)
    rows = []
    for a in D.active:
        id_ = a["AccountID"]
        au = usage.get((id_, m[-1]), {}).get("ActiveUsers", 0)
        util = excel_round(au / a["LicensedSeats"], 4)
        l3 = sum(usage[(id_, x)]["Logins"] for x in m[-3:] if (id_, x) in usage)
        p3 = sum(usage[(id_, x)]["Logins"] for x in m[-6:-3] if (id_, x) in usage)
        trend = excel_round(l3 / p3 - 1, 4) if p3 else 0
        tk = sum(1 for t in D.tickets if t["AccountID"] == id_ and t["ClosedDate"] is None
                 and t["Priority"] in ("P1", "P2"))
        sc = [n["Score"] for n in D.nps if n["AccountID"] == id_ and n["ResponseDate"] > cut]
        nps = sum(sc) / len(sc) if sc else ""
        days = (a["RenewalDate"] - D.as_of).days
        comps = [s_adoption(util), s_trend(trend), s_support(tk), s_nps(nps)]
        total = 0.0
        for s, w in zip(comps, WEIGHTS):
            total += s * w
        hs = excel_round(total, 1)
        cat = "Green" if hs >= 75 else "Yellow" if hs >= 50 else "Red"
        rows.append(dict(acc=a, au=au, util=util, l3=l3, p3=p3, trend=trend, tk=tk, nps=nps, days=days,
                         comps=comps, hs=hs, cat=cat,
                         flag="⚠ Riesgo renovación" if cat == "Red" and days <= 180 else ""))
    for r in rows:
        r["rank"] = 1 + sum(1 for o in rows if o["hs"] < r["hs"])
    D.health = rows
    return rows


def sheet_e3(bk, D):
    ws = bk.ws["E3_Health_Score"]
    H = D.health
    bk.title(ws, "E3 · Customer Health Score",
             "Objetivo: calcular para cada cuenta activa un health score de 0-100 combinando adopción (uso de "
             "licencias), tendencia de uso, tickets críticos abiertos y NPS; clasificar Green/Yellow/Red y "
             "marcar las cuentas en rojo que renuevan en ≤180 días. Los pesos están en la fila 4 (deben sumar 100%).",
             "Contexto: es lo que pide la vacante — 'build and refine health-score and KPI frameworks'. En la "
             "entrevista explica POR QUÉ elegiste cada métrica y peso, y qué validarías con los CSMs "
             "(¿las cuentas Red realmente están en riesgo?).", width=21)
    A, U, T, N = T_ACC, T_USE, T_TCK, T_NPS
    labels = ["AccountID", "AccountName", "Segment", "CSM_ID", "ARR_USD", "LicensedSeats",
              "ActiveUsers (LastMonth)", "Seat Utilization", "Logins últimos 3m", "Logins 3m previos",
              "Tendencia uso %", "Tickets P1/P2 abiertos", "NPS prom. 12m", "Días a renovación",
              "Score Adopción", "Score Tendencia", "Score Soporte", "Score NPS", "HEALTH SCORE",
              "Categoría", "Alerta", "Rank (1 = peor)"]
    hints = ["Dado", "XLOOKUP", "XLOOKUP", "XLOOKUP", "XLOOKUP", "XLOOKUP",
             "SUMIFS(tbl_Usage[ActiveUsers], ...[AccountID], A8, ...[Month], LastMonth)",
             "IFERROR(ROUND(G8/F8, 4), 0)",
             "SUMIFS(...[Logins], ..., [Month], \">=\"&EDATE(LastMonth,-2))",
             "Meses −5 a −3: \">=\"&EDATE(LastMonth,-5) y \"<\"&EDATE(LastMonth,-2)",
             "IFERROR(ROUND(I8/J8-1, 4), 0)",
             "SUM(COUNTIFS(...[AccountID],A8, ...[ClosedDate],\"\", ...[Priority],{\"P1\",\"P2\"}))",
             "IFERROR(AVERAGEIFS(tbl_NPS[Score], ..., [ResponseDate], \">\"&EDATE(AsOfDate,-12)), \"\")",
             "XLOOKUP(RenewalDate) − AsOfDate",
             "IFS: ≥80%→100, ≥60%→75, ≥40%→50, ≥20%→25, si no 0",
             "IFS: ≥10%→100, ≥0%→75, ≥−10%→50, ≥−25%→25, si no 0",
             "IFS: 0→100, 1→60, ≤3→30, si no 0",
             "Sin NPS (\"\")→50; IFS: ≥9→100, ≥7→70, ≥5→40, si no 10",
             "ROUND(SUMPRODUCT(O8:R8, $O$4:$R$4), 1)",
             "IFS: ≥75 Green, ≥50 Yellow, si no Red",
             "Red y renovación ≤180 días → \"⚠ Riesgo renovación\", si no \"\"",
             "RANK.EQ(S8, $S$8:$S$n, 1)"]
    # pesos
    ws.write(3, 13, "Pesos →", bk.f(bold=True, align="right"))
    for i, w in enumerate(WEIGHTS):
        ws.write_number(3, 14 + i, w, bk.f(num_format="0%", bold=True, bg_color="#DDEBF7", border=1))
    ws.write_formula(3, 18, "=SUM(O4:R4)", bk.f(num_format="0%", bold=True), 1)
    ws.write_formula(3, 19, '=IF(ROUND(S4,4)=1,"Pesos OK","Los pesos deben sumar 100%")',
                     bk.f(italic=True, font_color="#C00000"), "Pesos OK")
    ws.data_validation(3, 14, 3, 17, {"validate": "decimal", "criteria": "between", "minimum": 0,
                                      "maximum": 1, "error_message": "Usa un % entre 0% y 100%"})
    bk.header(ws, 5, 0, labels, hints)
    ws.set_row(5, 30)
    ws.set_row(6, 75)
    chk_c = 23
    chk_cols = [7, 10, 11, 12, 18, 19, 21]
    bk.header(ws, 5, chk_c, [f"✓{xl_col_to_name(c)}" for c in chk_cols])
    ws.set_column(0, 0, 10)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 21, 10.5)
    ws.set_column(18, 18, 11, bk.f(bold=True))
    ws.set_column(20, 20, 18)
    ws.set_column(22, 22, 2)
    ws.set_column(chk_c, chk_c + len(chk_cols) - 1, 4)

    first = 7
    for i, h in enumerate(H):
        r = first + i
        R = r + 1
        a = h["acc"]
        x = f"$A{R}"
        bk.given(ws, r, 0, a["AccountID"])
        bk.answer(ws, r, 1, bk.XL(x, col(A, "AccountID"), col(A, "AccountName")))
        bk.answer(ws, r, 2, bk.XL(x, col(A, "AccountID"), col(A, "Segment")))
        bk.answer(ws, r, 3, bk.XL(x, col(A, "AccountID"), col(A, "CSM_ID")))
        bk.answer(ws, r, 4, bk.XL(x, col(A, "AccountID"), col(A, "ARR_USD")), nf=NF_USD)
        bk.answer(ws, r, 5, bk.XL(x, col(A, "AccountID"), col(A, "LicensedSeats")), nf=NF_INT)
        bk.answer(ws, r, 6, f'SUMIFS({col(U, "ActiveUsers")},{col(U, "AccountID")},{x},{col(U, "Month")},LastMonth)',
                  nf=NF_INT)
        bk.answer(ws, r, 7, f"IFERROR(ROUND(G{R}/F{R},4),0)", nf=NF_PCT, expected=h["util"],
                  chk=(r, chk_c), tol=0.00005)
        bk.answer(ws, r, 8, f'SUMIFS({col(U, "Logins")},{col(U, "AccountID")},{x},{col(U, "Month")},'
                            f'">="&EDATE(LastMonth,-2))', nf=NF_INT)
        bk.answer(ws, r, 9, f'SUMIFS({col(U, "Logins")},{col(U, "AccountID")},{x},{col(U, "Month")},'
                            f'">="&EDATE(LastMonth,-5),{col(U, "Month")},"<"&EDATE(LastMonth,-2))', nf=NF_INT)
        bk.answer(ws, r, 10, f"IFERROR(ROUND(I{R}/J{R}-1,4),0)", nf=NF_PCT, expected=h["trend"],
                  chk=(r, chk_c + 1), tol=0.00005)
        bk.answer(ws, r, 11, f'SUM(COUNTIFS({col(T, "AccountID")},{x},{col(T, "ClosedDate")},"",'
                             f'{col(T, "Priority")},{{"P1","P2"}}))', expected=h["tk"], chk=(r, chk_c + 2))
        bk.answer(ws, r, 12, f'IFERROR(AVERAGEIFS({col(N, "Score")},{col(N, "AccountID")},{x},'
                             f'{col(N, "ResponseDate")},">"&EDATE(AsOfDate,-12)),"")', nf=NF_DEC,
                  expected=h["nps"], chk=(r, chk_c + 3), tol=0.005)
        bk.answer(ws, r, 13, f'{bk.XL(x, col(A, "AccountID"), col(A, "RenewalDate"))}-AsOfDate', nf=NF_INT)
        bk.answer(ws, r, 14, f"IFS(H{R}>=0.8,100,H{R}>=0.6,75,H{R}>=0.4,50,H{R}>=0.2,25,TRUE,0)")
        bk.answer(ws, r, 15, f"IFS(K{R}>=0.1,100,K{R}>=0,75,K{R}>=-0.1,50,K{R}>=-0.25,25,TRUE,0)")
        bk.answer(ws, r, 16, f"IFS(L{R}=0,100,L{R}=1,60,L{R}<=3,30,TRUE,0)")
        bk.answer(ws, r, 17, f'IF(M{R}="",50,IFS(M{R}>=9,100,M{R}>=7,70,M{R}>=5,40,TRUE,10))')
        bk.answer(ws, r, 18, f"ROUND(SUMPRODUCT(O{R}:R{R},$O$4:$R$4),1)", nf=NF_DEC, expected=h["hs"],
                  chk=(r, chk_c + 4), tol=0.01, bold=True)
        bk.answer(ws, r, 19, f'IFS(S{R}>=75,"Green",S{R}>=50,"Yellow",TRUE,"Red")', expected=h["cat"],
                  chk=(r, chk_c + 5))
        bk.answer(ws, r, 20, f'IF(AND(T{R}="Red",N{R}<=180),"⚠ Riesgo renovación","")')
        bk.answer(ws, r, 21, f"RANK.EQ(S{R},$S${first + 1}:$S${first + len(H)},1)", expected=h["rank"],
                  chk=(r, chk_c + 6))
    last = first + len(H) - 1
    L = last + 1
    D.e3_rows = (first + 1, L)
    bk.check_cf(ws, f"X{first + 1}:AD{L}")
    rng_cat = f"T{first + 1}:T{L}"
    ws.conditional_format(rng_cat, {"type": "cell", "criteria": "==", "value": '"Green"',
                                    "format": bk.f(bg_color="#C6EFCE", font_color="#006100")})
    ws.conditional_format(rng_cat, {"type": "cell", "criteria": "==", "value": '"Yellow"',
                                    "format": bk.f(bg_color="#FFEB9C", font_color="#9C5700")})
    ws.conditional_format(rng_cat, {"type": "cell", "criteria": "==", "value": '"Red"',
                                    "format": bk.f(bg_color="#FFC7CE", font_color="#9C0006")})
    ws.conditional_format(f"S{first + 1}:S{L}", {"type": "data_bar", "bar_color": "#5B9BD5",
                                                 "min_type": "num", "min_value": 0,
                                                 "max_type": "num", "max_value": 100})
    ws.conditional_format(f"U{first + 1}:U{L}", {"type": "text", "criteria": "containing", "value": "Riesgo",
                                                 "format": bk.f(bold=True, font_color="#C00000")})
    ws.freeze_panes(first, 2)

    # nombres definidos para usar en otras hojas
    for name, c in [("HS_ID", "A"), ("HS_Name", "B"), ("HS_Segment", "C"), ("HS_CSM", "D"), ("HS_ARR", "E"),
                    ("HS_Tickets", "L"), ("HS_Days", "N"), ("HS_Score", "S"), ("HS_Cat", "T")]:
        bk.wb.define_name(name, f"=E3_Health_Score!${c}${first + 1}:${c}${L}")

    # resumen
    sc = 31
    ws.set_column(sc - 1, sc - 1, 2)
    ws.set_column(sc, sc, 26)
    ws.set_column(sc + 1, sc + 2, 14)
    ws.set_column(sc + 3, sc + 3, 4)
    bk.note(ws, 3, sc, "Resumen (usa los nombres definidos HS_Cat, HS_ARR, ...)", bold=True, font_color=C_TITLE)
    bk.header(ws, 5, sc, ["Categoría", "# Cuentas", "ARR", "Chk"])
    cnt = Counter(h["cat"] for h in H)
    arr = defaultdict(float)
    for h in H:
        arr[h["cat"]] += h["acc"]["ARR_USD"]
    for i, cat in enumerate(["Green", "Yellow", "Red"]):
        r = 6 + i
        bk.given(ws, r, sc, cat, bold=True)
        bk.answer(ws, r, sc + 1, f'COUNTIF(HS_Cat,{xl_rowcol_to_cell(r, sc)})', expected=cnt[cat], chk=(r, sc + 3))
        bk.answer(ws, r, sc + 2, f'SUMIFS(HS_ARR,HS_Cat,{xl_rowcol_to_cell(r, sc)})', nf=NF_USD,
                  expected=arr[cat], chk=(r + 9, sc + 3))
    r = 10
    items = [("Health score promedio", "AVERAGE(HS_Score)", sum(h["hs"] for h in H) / len(H), NF_DEC),
             ("ARR ponderado por health (prom.)", "SUMPRODUCT(HS_Score,HS_ARR)/SUM(HS_ARR)",
              sum(h["hs"] * h["acc"]["ARR_USD"] for h in H) / sum(h["acc"]["ARR_USD"] for h in H), NF_DEC),
             ("# Cuentas con alerta de renovación", 'COUNTIFS(HS_Cat,"Red",HS_Days,"<=180")',
              sum(1 for h in H if h["flag"]), NF_INT),
             ("ARR en riesgo (Red, ≤180 días)", 'SUMIFS(HS_ARR,HS_Cat,"Red",HS_Days,"<=180")',
              sum(h["acc"]["ARR_USD"] for h in H if h["flag"]), NF_USD)]
    for i, (lab, fml, exp, nf) in enumerate(items):
        rr = r + i
        bk.given(ws, rr, sc, lab, bold=True)
        bk.answer(ws, rr, sc + 2, fml, nf=nf, expected=exp, chk=(rr, sc + 3), tol=0.05)
    ws.write(15, sc, "✓ de la columna ARR (G/Y/R) →", bk.f(italic=True, font_size=8, align="right"))
    bk.check_cf(ws, f"AI7:AI18")
    chart = bk.wb.add_chart({"type": "doughnut"})
    chart.add_series({"name": "Cuentas por categoría",
                      "categories": ["E3_Health_Score", 6, sc, 8, sc],
                      "values": ["E3_Health_Score", 6, sc + 1, 8, sc + 1],
                      "points": [{"fill": {"color": "#70AD47"}}, {"fill": {"color": "#FFC000"}},
                                 {"fill": {"color": "#C00000"}}]})
    chart.set_title({"name": "Cuentas por categoría"})
    chart.set_size({"width": 360, "height": 240})
    ws.insert_chart(19, sc, chart)


# ---------------------------------------------------------------------------
# E4 - Auditoría de reportes
# ---------------------------------------------------------------------------

def recommend(rep, v90):
    if rep["Type"] == "Contractual":
        return "Mantener (contractual)"
    if v90 == 0:
        return "Retirar"
    if v90 < 10:
        return "Consolidar / revisar"
    if rep["Frequency"] == "Ad hoc" and v90 >= 30:
        return "Automatizar"
    return "Mantener"


RECS = ["Retirar", "Consolidar / revisar", "Automatizar", "Mantener", "Mantener (contractual)"]


def sheet_e4(bk, D):
    ws = bk.ws["E4_Auditoria_Reportes"]
    bk.title(ws, "E4 · Auditoría del uso de reportes",
             "Objetivo: con el log de uso (Report_Usage_Log) y el inventario (Report_Catalog) mide cuánto se usa "
             "realmente cada reporte, cuánto cuesta mantenerlo y genera una recomendación: Retirar, Consolidar, "
             "Automatizar o Mantener. Luego resume cuántas horas se liberarían.",
             "Contexto: la vacante pide 'analizar y auditar' los reportes actuales para saber qué mejorar, quitar o "
             "agregar. Fíjate en reportes duplicados (Login v1 vs v2, 'Copy of...'), reportes ad hoc que se piden "
             "tanto que conviene automatizarlos y reportes contractuales que nadie abre (¿se envían por otro lado?).",
             width=20)
    L, R_ = T_LOG, T_REP
    labels = ["ReportID", "ReportName", "Type", "Frequency", "Views (180d)", "Exports (180d)", "Usuarios únicos",
              "Último uso", "Días sin uso", "Views últimos 90d", "Horas mant./mes", "Costo por vista (USD)",
              "Recomendación"]
    hints = ["Dado", "XLOOKUP", "XLOOKUP", "XLOOKUP",
             "COUNTIFS(tbl_ReportLog[ReportID], A7, ...[Action], \"View\")", "Igual con \"Export\"",
             "IFERROR(ROWS(UNIQUE(FILTER(tbl_ReportLog[UserEmail], tbl_ReportLog[ReportID]=A7))), 0)",
             "Si no hay registros → \"\"; si no MAXIFS(tbl_ReportLog[Date], ...)",
             "IF(H7=\"\", \"\", AsOfDate − H7)",
             "COUNTIFS(..., \"View\", ...[Date], \">\"&AsOfDate-90)", "XLOOKUP MaintHrsPerMonth",
             "IFERROR(K7*6*CostPerHour/E7, \"\")",
             "IFS: Contractual→\"Mantener (contractual)\"; J=0→\"Retirar\"; J<10→\"Consolidar / revisar\"; "
             "Ad hoc y J≥30→\"Automatizar\"; si no \"Mantener\""]
    bk.header(ws, 4, 0, labels, hints)
    ws.set_row(5, 85)
    chk_c = 14
    chk_cols = [4, 5, 6, 7, 9, 11, 12]
    bk.header(ws, 4, chk_c, [f"✓{xl_col_to_name(c)}" for c in chk_cols])
    ws.set_column(0, 0, 9)
    ws.set_column(1, 1, 32)
    ws.set_column(2, 11, 11)
    ws.set_column(12, 12, 22)
    ws.set_column(13, 13, 2)
    ws.set_column(chk_c, chk_c + 6, 4)

    stats = {}
    window90 = D.as_of - dt.timedelta(days=90)
    for rep in D.reports:
        logs = [x for x in D.log if x["ReportID"] == rep["ReportID"]]
        views = sum(1 for x in logs if x["Action"] == "View")
        v90 = sum(1 for x in logs if x["Action"] == "View" and x["Date"] > window90)
        last = max((x["Date"] for x in logs), default=None)
        stats[rep["ReportID"]] = dict(
            views=views, exports=sum(1 for x in logs if x["Action"] == "Export"),
            users=len({x["UserEmail"] for x in logs}), last=last, v90=v90,
            cpv=(rep["MaintHrsPerMonth"] * 6 * 35 / views) if views else "",
            rec=recommend(rep, v90))
    D.report_stats = stats

    first = 6
    for i, rep in enumerate(D.reports):
        r = first + i
        Rn = r + 1
        x = f"$A{Rn}"
        s = stats[rep["ReportID"]]
        bk.given(ws, r, 0, rep["ReportID"])
        bk.answer(ws, r, 1, bk.XL(x, col(R_, "ReportID"), col(R_, "ReportName")))
        bk.answer(ws, r, 2, bk.XL(x, col(R_, "ReportID"), col(R_, "Type")))
        bk.answer(ws, r, 3, bk.XL(x, col(R_, "ReportID"), col(R_, "Frequency")))
        bk.answer(ws, r, 4, f'COUNTIFS({col(L, "ReportID")},{x},{col(L, "Action")},"View")', nf=NF_INT,
                  expected=s["views"], chk=(r, chk_c))
        bk.answer(ws, r, 5, f'COUNTIFS({col(L, "ReportID")},{x},{col(L, "Action")},"Export")', nf=NF_INT,
                  expected=s["exports"], chk=(r, chk_c + 1))
        bk.answer(ws, r, 6, f'IFERROR(ROWS(UNIQUE(FILTER({col(L, "UserEmail")},{col(L, "ReportID")}={x}))),0)',
                  expected=s["users"], chk=(r, chk_c + 2))
        bk.answer(ws, r, 7, f'IF(COUNTIF({col(L, "ReportID")},{x})=0,"",MAXIFS({col(L, "Date")},'
                            f'{col(L, "ReportID")},{x}))', nf=NF_DATE, expected=s["last"] or "",
                  chk=(r, chk_c + 3))
        bk.answer(ws, r, 8, f'IF(H{Rn}="","",AsOfDate-H{Rn})', nf=NF_INT)
        bk.answer(ws, r, 9, f'COUNTIFS({col(L, "ReportID")},{x},{col(L, "Action")},"View",{col(L, "Date")},'
                            f'">"&AsOfDate-90)', nf=NF_INT, expected=s["v90"], chk=(r, chk_c + 4))
        bk.answer(ws, r, 10, bk.XL(x, col(R_, "ReportID"), col(R_, "MaintHrsPerMonth")), nf=NF_DEC)
        bk.answer(ws, r, 11, f'IFERROR(K{Rn}*6*CostPerHour/E{Rn},"")', nf=NF_USD2, expected=s["cpv"],
                  chk=(r, chk_c + 5), tol=0.01)
        bk.answer(ws, r, 12, f'IFS(C{Rn}="Contractual","Mantener (contractual)",J{Rn}=0,"Retirar",'
                             f'J{Rn}<10,"Consolidar / revisar",AND(D{Rn}="Ad hoc",J{Rn}>=30),"Automatizar",'
                             f'TRUE,"Mantener")', expected=s["rec"], chk=(r, chk_c + 6), bold=True)
    last = first + len(D.reports)
    bk.check_cf(ws, f"O{first + 1}:U{last}")
    colors = {"Retirar": ("#FFC7CE", "#9C0006"), "Consolidar / revisar": ("#FFEB9C", "#9C5700"),
              "Automatizar": ("#DDEBF7", "#1F4E78"), "Mantener": ("#C6EFCE", "#006100"),
              "Mantener (contractual)": ("#E4DFEC", "#5B2C6F")}
    for rec, (bg, fc) in colors.items():
        ws.conditional_format(f"M{first + 1}:M{last}", {"type": "cell", "criteria": "==", "value": f'"{rec}"',
                                                        "format": bk.f(bg_color=bg, font_color=fc, bold=True)})
    ws.conditional_format(f"E{first + 1}:E{last}", {"type": "data_bar", "bar_color": "#5B9BD5"})
    ws.freeze_panes(first, 2)

    # resumen
    sc = 22
    ws.set_column(sc - 1, sc - 1, 2)
    ws.set_column(sc, sc, 30)
    ws.set_column(sc + 1, sc + 2, 13)
    ws.set_column(sc + 3, sc + 3, 4)
    bk.note(ws, 3, sc, "Resumen para liderazgo", bold=True, font_color=C_TITLE)
    bk.header(ws, 4, sc, ["Recomendación", "# Reportes", "Horas/mes", "Chk"],
              ["", "COUNTIF($M$7:$M$n, W7)", "SUMIF(..., $K$7:$K$n)", ""])
    for i, rec in enumerate(RECS):
        r = 6 + i
        reps = [rp for rp in D.reports if stats[rp["ReportID"]]["rec"] == rec]
        bk.given(ws, r, sc, rec, bold=True)
        cell = xl_rowcol_to_cell(r, sc)
        bk.answer(ws, r, sc + 1, f"COUNTIF($M${first + 1}:$M${last},{cell})", expected=len(reps), chk=(r, sc + 3))
        bk.answer(ws, r, sc + 2, f"SUMIF($M${first + 1}:$M${last},{cell},$K${first + 1}:$K${last})", nf=NF_DEC,
                  expected=sum(rp["MaintHrsPerMonth"] for rp in reps), chk=(r + 6, sc + 3))
    ws.write(11, sc, "✓ de Horas/mes →", bk.f(italic=True, font_size=8, align="right"))
    retire = [rp for rp in D.reports if stats[rp["ReportID"]]["rec"] == "Retirar"]
    r = 18
    bk.given(ws, r, sc, "Ahorro anual si se retiran (USD)", bold=True)
    bk.answer(ws, r, sc + 2, f'SUMIF($M${first + 1}:$M${last},"Retirar",$K${first + 1}:$K${last})*12*CostPerHour',
              nf=NF_USD, expected=sum(rp["MaintHrsPerMonth"] for rp in retire) * 12 * 35, chk=(r, sc + 3))
    bk.given(ws, r + 1, sc, "% reportes sin views en 90 días", bold=True)
    bk.answer(ws, r + 1, sc + 2, f"COUNTIF($J${first + 1}:$J${last},0)/COUNTA($A${first + 1}:$A${last})", nf=NF_PCT,
              expected=sum(1 for s in stats.values() if s["v90"] == 0) / len(stats), chk=(r + 1, sc + 3), tol=0.0005)

    r = 22
    bk.note(ws, r - 1, sc, "¿Quién consume los reportes? (views por rol)", bold=True, font_color=C_TITLE)
    bk.header(ws, r, sc, ["UserRole", "Views", "% del total", "Chk"],
              ["", "COUNTIFS(...[UserRole], W24, ...[Action], \"View\")", "X / SUM(X)", ""])
    tot = sum(1 for x in D.log if x["Action"] == "View")
    for i, role in enumerate(ROLES):
        rr = r + 2 + i
        v = sum(1 for x in D.log if x["UserRole"] == role and x["Action"] == "View")
        bk.given(ws, rr, sc, role)
        bk.answer(ws, rr, sc + 1, f'COUNTIFS({col(L, "UserRole")},{xl_rowcol_to_cell(rr, sc)},{col(L, "Action")},"View")',
                  nf=NF_INT, expected=v, chk=(rr, sc + 3))
        bk.answer(ws, rr, sc + 2, f"{xl_rowcol_to_cell(rr, sc + 1)}/SUM(${xl_col_to_name(sc + 1)}${r + 3}:"
                                  f"${xl_col_to_name(sc + 1)}${r + 2 + len(ROLES)})", nf=NF_PCT)
    bk.check_cf(ws, f"Z7:Z{r + 2 + len(ROLES)}")
    bk.note(ws, r + 3 + len(ROLES), sc,
            "Para la entrevista: redacta 3-5 bullets para liderazgo con los hallazgos (p. ej. 'X reportes sin uso "
            "en 90 días → retirarlos libera N h/mes'; 'Login v1 y v2 duplican información'; 'el reporte ad hoc Y "
            "se pide cada semana → automatizar'; 'un reporte contractual sin views: validar con el CSM que se "
            "esté entregando al cliente').", merge_to=sc + 3, rows=7, italic=True)


# ---------------------------------------------------------------------------
# E5 - Limpieza de datos
# ---------------------------------------------------------------------------

SEG_MAP = {"ENTERPRISE": "Enterprise", "ENT": "Enterprise", "MID-MARKET": "Mid-Market", "MID MARKET": "Mid-Market",
           "MM": "Mid-Market", "SMB": "SMB", "SMALL BUSINESS": "SMB"}


def parse_raw_date(s):
    if s[4] == "-":
        return dt.date(int(s[:4]), int(s[5:7]), int(s[-2:]))
    return dt.date(int(s[-4:]), int(s[3:5]), int(s[:2]))


def sheet_e5(bk, D):
    ws = bk.ws["E5_Limpieza_Datos"]
    bk.title(ws, "E5 · Limpieza y auditoría de datos (export del CRM)",
             "Objetivo: un CSM te pasó este export del CRM (columnas A-F, tal cual). Límpialo con fórmulas "
             "(columnas H-M), detecta duplicados, cuentas que no existen en el maestro y diferencias de ARR "
             "(columnas N-Q). Después resume la calidad de los datos a la derecha.",
             "Contexto: antes de cualquier reporte o health score hay que confiar en los datos. En la entrevista "
             "describe tu proceso: estandarizar → deduplicar → conciliar contra la fuente de verdad → documentar "
             "y corregir en el origen (no solo en Excel).", width=27)
    labels = ["Account ID (raw)", "Account Name (raw)", "Segment (raw)", "ARR (raw)", "Renewal (raw)",
              "CSM Email (raw)", "", "ID limpio", "Nombre limpio", "Segment limpio", "ARR (número)",
              "Renewal (fecha)", "Email limpio", "¿Duplicado?", "¿Existe en maestro?", "Dif. ARR vs maestro",
              "# Problemas"]
    hints = ["", "", "", "", "Formatos: AAAA-MM-DD o DD/MM/AAAA", "", "",
             "UPPER(TRIM(A8))", "PROPER(TRIM(B8))",
             "SWITCH(UPPER(TRIM(C8)), \"ENTERPRISE\",\"Enterprise\", \"ENT\",\"Enterprise\", \"MID-MARKET\",..., "
             "\"MID MARKET\",..., \"MM\",..., \"SMB\",\"SMB\", \"SMALL BUSINESS\",\"SMB\", \"Revisar\")",
             "NUMBERVALUE(SUBSTITUTE(SUBSTITUTE(D8,\"$\",\"\"),\"USD\",\"\"), \".\", \",\")",
             "IF(MID(E8,5,1)=\"-\", DATE(LEFT(E8,4),MID(E8,6,2),RIGHT(E8,2)), DATE(RIGHT(E8,4),MID(E8,4,2),LEFT(E8,2)))",
             "LOWER(TRIM(F8))", "IF(COUNTIF($H$8:H8, H8)>1, \"Duplicado\", \"\")  ← rango que crece",
             "IF(ISNUMBER(MATCH(H8, tbl_Accounts[AccountID], 0)), \"Sí\", \"No\")",
             "IFERROR(K8 − XLOOKUP(H8, ...[AccountID], ...[ARR_USD]), \"\")",
             "(N8<>\"\") + (O8=\"No\") + IF(ISNUMBER(P8), --(P8<>0), 0) + (M8=\"\")"]
    bk.header(ws, 5, 0, labels, hints)
    ws.set_row(6, 95)
    ws.set_column(0, 0, 13)
    ws.set_column(1, 1, 24)
    ws.set_column(2, 4, 13)
    ws.set_column(5, 5, 28)
    ws.set_column(6, 6, 2)
    ws.set_column(7, 7, 11)
    ws.set_column(8, 8, 22)
    ws.set_column(9, 11, 12)
    ws.set_column(12, 12, 28)
    ws.set_column(13, 16, 11)
    ws.set_column(17, 17, 2)
    chk_c = 18
    ws.set_column(chk_c, chk_c + 9, 4)
    bk.header(ws, 5, chk_c, [f"✓{xl_col_to_name(c)}" for c in range(7, 17)])
    first = 7
    seen = Counter()
    exp_rows = []
    for i, raw in enumerate(D.raw):
        r = first + i
        R = r + 1
        for c, v in enumerate(raw):
            ws.write_string(r, c, v, bk.f(border=1, border_color="#BFBFBF", bg_color="#FBE5D6"))
        cid = excel_trim(raw[0]).upper()
        name = excel_proper(excel_trim(raw[1]))
        seg = SEG_MAP.get(excel_trim(raw[2]).upper(), "Revisar")
        arr = float(raw[3].replace("$", "").replace("USD", "").replace(",", "").strip())
        ren = parse_raw_date(raw[4])
        em = excel_trim(raw[5]).lower()
        seen[cid] += 1
        dup = "Duplicado" if seen[cid] > 1 else ""
        acc = D.acc_by_id.get(cid)
        exists = "Sí" if acc else "No"
        diff = arr - acc["ARR_USD"] if acc else ""
        issues = (dup != "") + (exists == "No") + (1 if diff not in ("", 0) else 0) + (em == "")
        exp_rows.append(dict(id=cid, dup=dup, exists=exists, diff=diff, issues=issues, em=em))
        exps = [cid, name, seg, arr, ren, em, dup, exists, diff, issues]
        fmls = [f"UPPER(TRIM(A{R}))", f"PROPER(TRIM(B{R}))",
                f'SWITCH(UPPER(TRIM(C{R})),"ENTERPRISE","Enterprise","ENT","Enterprise","MID-MARKET","Mid-Market",'
                f'"MID MARKET","Mid-Market","MM","Mid-Market","SMB","SMB","SMALL BUSINESS","SMB","Revisar")',
                f'NUMBERVALUE(SUBSTITUTE(SUBSTITUTE(D{R},"$",""),"USD",""),".",",")',
                f'IF(MID(E{R},5,1)="-",DATE(LEFT(E{R},4),MID(E{R},6,2),RIGHT(E{R},2)),'
                f'DATE(RIGHT(E{R},4),MID(E{R},4,2),LEFT(E{R},2)))',
                f"LOWER(TRIM(F{R}))",
                f'IF(COUNTIF($H${first + 1}:H{R},H{R})>1,"Duplicado","")',
                f'IF(ISNUMBER(MATCH(H{R},{col(T_ACC, "AccountID")},0)),"Sí","No")',
                f'IFERROR(K{R}-{bk.XL(f"H{R}", col(T_ACC, "AccountID"), col(T_ACC, "ARR_USD"))},"")',
                f'(N{R}<>"")+(O{R}="No")+IF(ISNUMBER(P{R}),--(P{R}<>0),0)+(M{R}="")']
        nfs = [None, None, None, NF_USD, NF_DATE, None, None, None, NF_USD, None]
        for j in range(10):
            bk.answer(ws, r, 7 + j, fmls[j], nf=nfs[j], expected=exps[j], chk=(r, chk_c + j), tol=0.01)
    last = first + len(D.raw)
    bk.check_cf(ws, f"S{first + 1}:AB{last}")
    for rng_, val in [(f"N{first + 1}:N{last}", '"Duplicado"'), (f"O{first + 1}:O{last}", '"No"')]:
        ws.conditional_format(rng_, {"type": "cell", "criteria": "==", "value": val,
                                     "format": bk.f(bg_color="#FFC7CE", font_color="#9C0006")})
    ws.conditional_format(f"Q{first + 1}:Q{last}", {"type": "cell", "criteria": ">", "value": 0,
                                                    "format": bk.f(bg_color="#FFC7CE", font_color="#9C0006")})
    ws.conditional_format(f"J{first + 1}:J{last}", {"type": "cell", "criteria": "==", "value": '"Revisar"',
                                                    "format": bk.f(bg_color="#FFEB9C")})
    ws.freeze_panes(first, 1)

    sc = 29
    ws.set_column(sc - 1, sc - 1, 2)
    ws.set_column(sc, sc, 34)
    ws.set_column(sc + 1, sc + 1, 12)
    ws.set_column(sc + 2, sc + 2, 4)
    bk.note(ws, 4, sc, "Resumen de calidad de datos", bold=True, font_color=C_TITLE)
    bk.header(ws, 5, sc, ["Métrica", "Valor", "Chk"])
    n = len(D.raw)
    rngH, rngN, rngO, rngP = (f"$H${first + 1}:$H${last}", f"$N${first + 1}:$N${last}",
                              f"$O${first + 1}:$O${last}", f"$P${first + 1}:$P${last}")
    items = [
        ("Filas en el export", f"COUNTA($A${first + 1}:$A${last})", n, NF_INT),
        ("IDs únicos", f"SUMPRODUCT(1/COUNTIF({rngH},{rngH}))", len({e["id"] for e in exp_rows}), NF_INT),
        ("Filas duplicadas", f'COUNTIF({rngN},"Duplicado")', sum(1 for e in exp_rows if e["dup"]), NF_INT),
        ("IDs que no existen en el maestro", f'COUNTIF({rngO},"No")',
         sum(1 for e in exp_rows if e["exists"] == "No"), NF_INT),
        ("Filas con ARR distinto al maestro", f"SUMPRODUCT(--ISNUMBER({rngP}),--({rngP}<>0))",
         sum(1 for e in exp_rows if e["diff"] not in ("", 0)), NF_INT),
        ("Filas sin email de CSM", f"COUNTBLANK($M${first + 1}:$M${last})", sum(1 for e in exp_rows if e["em"] == ""),
         NF_INT),
        ("% filas sin problemas", f"COUNTIF($Q${first + 1}:$Q${last},0)/COUNTA($A${first + 1}:$A${last})",
         sum(1 for e in exp_rows if e["issues"] == 0) / n, NF_PCT),
    ]
    for i, (lab, fml, exp, nf) in enumerate(items):
        r = 6 + i
        bk.given(ws, r, sc, lab, bold=True)
        bk.answer(ws, r, sc + 1, fml, nf=nf, expected=exp, chk=(r, sc + 2), tol=0.0005)
    bk.check_cf(ws, f"AF7:AF{6 + len(items)}")
    bk.note(ws, 7 + len(items), sc,
            "Siguiente paso real: documenta las reglas de limpieza, avisa al dueño del dato (Sales Ops / CSM) y "
            "corrige en Salesforce. Opcional: repite esta limpieza con Power Query (E10) y compara.",
            merge_to=sc + 2, rows=4, italic=True)


# ---------------------------------------------------------------------------
# E6 - Pivots
# ---------------------------------------------------------------------------

def sheet_e6(bk, D):
    ws = bk.ws["E6_Pivots"]
    bk.title(ws, "E6 · Tablas dinámicas (Pivot Tables)",
             "Objetivo: crea cada tabla dinámica en una hoja nueva (Insertar > Tabla dinámica, origen = la tabla "
             "indicada) y copia aquí los resultados (a mano o con GETPIVOTDATA). Las ✓ te dicen si tu pivot está "
             "bien. En la versión de soluciones verás cómo obtener lo mismo con fórmulas.",
             "Tip: usa siempre una Tabla de Excel como origen (se expande sola), cambia 'Resumir valores por' y "
             "'Mostrar valores como' (% del total), agrupa fechas por mes, agrega segmentaciones (slicers) y una "
             "escala de tiempo. Recuerda 'Actualizar' después de cambiar datos.", width=14)
    ws.set_column(0, 0, 24)
    ws.set_column(1, 8, 13)
    ws.set_column(9, 9, 2)
    ws.set_column(10, 17, 4)
    A, U, T, L, N = T_ACC, T_USE, T_TCK, T_LOG, T_NPS
    active = [a for a in D.accounts if a["Status"] == "Active"]
    r = 4

    def block_title(r, text, steps):
        ws.merge_range(r, 0, r, 8, text, bk.f(bold=True, font_color="white", bg_color=C_HDR))
        ws.merge_range(r + 1, 0, r + 1, 8, steps, bk.f(text_wrap=True, valign="top", bg_color=C_HINT,
                                                       font_size=9))
        ws.set_row(r + 1, 58)
        return r + 2

    # P1
    r = block_title(r, "P1 · Cuentas y ARR por Producto (solo activas)",
                    "Origen: tbl_Accounts. Filtros: Status = Active. Filas: Product. Valores: Recuento de AccountID, "
                    "Suma de ARR_USD y otra vez Suma de ARR_USD mostrada como '% del total general'. Extra: agrega "
                    "un campo calculado ARR por licencia = ARR_USD / LicensedSeats y un slicer por Region.")
    bk.header(ws, r, 0, ["Product", "# Cuentas", "ARR", "% del ARR"])
    bk.header(ws, r, 10, ["✓B", "✓C", "✓D"])
    tot = sum(a["ARR_USD"] for a in active)
    for i, p in enumerate(PRODUCTS):
        rr = r + 1 + i
        x = f"$A{rr + 1}"
        mine = [a for a in active if a["Product"] == p]
        bk.given(ws, rr, 0, p)
        bk.answer(ws, rr, 1, f'COUNTIFS({col(A, "Product")},{x},{col(A, "Status")},"Active")',
                  expected=len(mine), chk=(rr, 10))
        bk.answer(ws, rr, 2, f'SUMIFS({col(A, "ARR_USD")},{col(A, "Product")},{x},{col(A, "Status")},"Active")',
                  nf=NF_USD, expected=sum(a["ARR_USD"] for a in mine), chk=(rr, 11))
        bk.answer(ws, rr, 3, f'C{rr + 1}/SUMIFS({col(A, "ARR_USD")},{col(A, "Status")},"Active")', nf=NF_PCT,
                  expected=sum(a["ARR_USD"] for a in mine) / tot, chk=(rr, 12), tol=0.0005)
    bk.check_cf(ws, f"K{r + 2}:M{r + 1 + len(PRODUCTS)}")
    r = r + len(PRODUCTS) + 2

    # P2
    r = block_title(r, "P2 · Usuarios activos por Mes y Segmento",
                    "Origen: tbl_Usage. Primero agrega a tbl_Usage una columna Segment con XLOOKUP hacia tbl_Accounts "
                    "(o usa el Modelo de datos con una relación por AccountID). Filas: Month (agrupado por meses/años). "
                    "Columnas: Segment. Valores: Suma de ActiveUsers. Extra: inserta un gráfico dinámico de líneas.")
    bk.header(ws, r, 0, ["Month"] + SEGMENTS)
    bk.header(ws, r, 10, ["✓B", "✓C", "✓D"])
    seg_of = {a["AccountID"]: a["Segment"] for a in D.accounts}
    agg = defaultdict(int)
    for u in D.usage:
        agg[(u["Month"], seg_of[u["AccountID"]])] += u["ActiveUsers"]
    for i, m in enumerate(D.months):
        rr = r + 1 + i
        bk.given(ws, rr, 0, m, NF_MONTH)
        for j, seg in enumerate(SEGMENTS):
            c = 1 + j
            fml = (f'SUMPRODUCT(SUMIFS({col(U, "ActiveUsers")},{col(U, "Month")},$A{rr + 1},{col(U, "AccountID")},'
                   f'{col(A, "AccountID")})*({col(A, "Segment")}={xl_col_to_name(c)}${r + 1}))')
            bk.answer(ws, rr, c, fml, nf=NF_INT, expected=agg[(m, seg)], chk=(rr, 10 + j))
    bk.check_cf(ws, f"K{r + 2}:M{r + 13}")
    r = r + 14

    # P3
    r = block_title(r, "P3 · Tickets por Categoría y Prioridad + CSAT",
                    "Origen: tbl_Tickets. Filas: Category. Columnas: Priority. Valores: Recuento de TicketID. Agrega "
                    "también Promedio de CSAT (Resumir valores por > Promedio) en una columna aparte. Extra: aplica "
                    "formato condicional de escala de color a la matriz.")
    bk.header(ws, r, 0, ["Category"] + PRIORITIES + ["CSAT promedio"])
    bk.header(ws, r, 10, ["✓B", "✓C", "✓D", "✓E", "✓F"])
    for i, cat in enumerate(TICKET_CATEGORIES):
        rr = r + 1 + i
        bk.given(ws, rr, 0, cat)
        for j, p in enumerate(PRIORITIES):
            c = 1 + j
            exp = sum(1 for t in D.tickets if t["Category"] == cat and t["Priority"] == p)
            bk.answer(ws, rr, c, f'COUNTIFS({col(T, "Category")},$A{rr + 1},{col(T, "Priority")},'
                                 f'{xl_col_to_name(c)}${r + 1})', expected=exp, chk=(rr, 10 + j))
        cs = [t["CSAT"] for t in D.tickets if t["Category"] == cat and t["CSAT"] is not None]
        bk.answer(ws, rr, 5, f'IFERROR(AVERAGEIFS({col(T, "CSAT")},{col(T, "Category")},$A{rr + 1}),"")',
                  nf=NF_DEC2, expected=sum(cs) / len(cs) if cs else "", chk=(rr, 14), tol=0.005)
    bk.check_cf(ws, f"K{r + 2}:O{r + 1 + len(TICKET_CATEGORIES)}")
    r = r + len(TICKET_CATEGORIES) + 2

    # P4
    r = block_title(r, "P4 · Top 5 reportes más vistos",
                    "Origen: tbl_ReportLog. Filtro: Action = View. Filas: ReportID. Valores: Recuento de LogID. "
                    "Ordena descendente y aplica Filtro de valor > Top 10 (cámbialo a 5). Extra: agrega ReportName "
                    "con el Modelo de datos o con una columna XLOOKUP en el log.")
    bk.header(ws, r, 0, ["Posición", "ReportID", "Views"])
    bk.header(ws, r, 10, ["✓B", "✓C"])
    views = Counter(x["ReportID"] for x in D.log if x["Action"] == "View")
    top = sorted(views.values(), reverse=True)[:5]
    for k in range(5):
        rr = r + 1 + k
        bk.given(ws, rr, 0, k + 1)
        bk.answer(ws, rr, 1, f'INDEX(SORTBY({col(T_REP, "ReportID")},COUNTIFS({col(L, "ReportID")},'
                             f'{col(T_REP, "ReportID")},{col(L, "Action")},"View"),-1),A{rr + 1})',
                  expected=top[k], chk=(rr, 10),
                  cmp=f'COUNTIFS({col(L, "ReportID")},{{ref}},{col(L, "Action")},"View")={{key}}')
        bk.answer(ws, rr, 2, f'SUMPRODUCT(LARGE(COUNTIFS({col(L, "ReportID")},{col(T_REP, "ReportID")},'
                             f'{col(L, "Action")},"View"),A{rr + 1}))', nf=NF_INT, expected=top[k], chk=(rr, 11))
    bk.check_cf(ws, f"K{r + 2}:L{r + 6}")
    r = r + 7

    # P5
    r = block_title(r, "P5 · Views por rol de usuario",
                    "Origen: tbl_ReportLog. Filtro: Action = View. Filas: UserRole. Valores: Recuento de LogID y "
                    "el mismo campo como '% del total de columnas'. Extra: Columnas = Date agrupada por Mes para "
                    "ver la tendencia; agrega una escala de tiempo (Timeline).")
    bk.header(ws, r, 0, ["UserRole", "Views"])
    bk.header(ws, r, 10, ["✓B"])
    for i, role in enumerate(ROLES):
        rr = r + 1 + i
        bk.given(ws, rr, 0, role)
        bk.answer(ws, rr, 1, f'COUNTIFS({col(L, "UserRole")},$A{rr + 1},{col(L, "Action")},"View")', nf=NF_INT,
                  expected=sum(1 for x in D.log if x["UserRole"] == role and x["Action"] == "View"), chk=(rr, 10))
    bk.check_cf(ws, f"K{r + 2}:K{r + 1 + len(ROLES)}")
    r = r + len(ROLES) + 2

    # P6
    r = block_title(r, "P6 · NPS por Segmento",
                    "Origen: tbl_NPS. Agrega 2 columnas a tbl_NPS: Segment (XLOOKUP a tbl_Accounts) y NPS_Category "
                    "=IFS([@Score]>=9,\"Promoter\",[@Score]>=7,\"Passive\",TRUE,\"Detractor\"). Pivot: Filas Segment, "
                    "Columnas NPS_Category, Valores Recuento de SurveyID. Luego calcula NPS = (P − D) / Total × 100 "
                    "fuera del pivot (usa GETPIVOTDATA).")
    bk.header(ws, r, 0, ["Segment", "Promoter", "Passive", "Detractor", "NPS"])
    bk.header(ws, r, 10, ["✓B", "✓C", "✓D", "✓E"])
    S, NA = col(N, "Score"), col(N, "AccountID")
    for i, seg in enumerate(SEGMENTS):
        rr = r + 1 + i
        R = rr + 1
        sc = [n["Score"] for n in D.nps if seg_of[n["AccountID"]] == seg]
        p = sum(1 for s in sc if s >= 9)
        pa = sum(1 for s in sc if 7 <= s <= 8)
        d = sum(1 for s in sc if s <= 6)
        bk.given(ws, rr, 0, seg)
        mask = f'*({col(A, "Segment")}=$A{R}))'
        bk.answer(ws, rr, 1, f'SUMPRODUCT(COUNTIFS({NA},{col(A, "AccountID")},{S},">=9"){mask}',
                  expected=p, chk=(rr, 10))
        bk.answer(ws, rr, 2, f'SUMPRODUCT(COUNTIFS({NA},{col(A, "AccountID")},{S},">=7",{S},"<=8"){mask}',
                  expected=pa, chk=(rr, 11))
        bk.answer(ws, rr, 3, f'SUMPRODUCT(COUNTIFS({NA},{col(A, "AccountID")},{S},"<=6"){mask}',
                  expected=d, chk=(rr, 12))
        bk.answer(ws, rr, 4, f"IFERROR((B{R}-D{R})/SUM(B{R}:D{R})*100,0)", nf=NF_DEC,
                  expected=(p - d) / (p + pa + d) * 100 if sc else 0, chk=(rr, 13), tol=0.05)
    bk.check_cf(ws, f"K{r + 2}:N{r + 4}")
    r += 5
    bk.note(ws, r, 0, "GETPIVOTDATA: escribe '=' y haz clic en una celda del pivot; Excel genera algo como "
                      "=GETPIVOTDATA(\"SurveyID\", Pivot!$A$3, \"Segment\", \"SMB\", \"NPS_Category\", \"Promoter\"). "
                      "Así tus cálculos no se rompen si el pivot cambia de tamaño.", merge_to=8, height=45, italic=True)


# ---------------------------------------------------------------------------
# E7 - Funciones de matriz dinámica
# ---------------------------------------------------------------------------

def sheet_e7(bk, D):
    ws = bk.ws["E7_Dinamicas_365"]
    bk.title(ws, "E7 · Funciones de matriz dinámica (Excel 365)",
             "Objetivo: resolver cada tarea con UNA fórmula en la celda amarilla; el resultado se 'derrama' (spill) "
             "hacia abajo/derecha. Deja libre el espacio gris debajo de cada celda. Solo algunas tareas tienen ✓.",
             "Si tu Excel no tiene estas funciones (2019 o anterior) verás #NAME?. Son muy valoradas porque "
             "reemplazan filtros manuales y copiar/pegar en reportes recurrentes.", width=12)
    ws.set_column(0, 0, 4)
    ws.set_column(1, 1, 16)
    ws.set_column(2, 12, 14)
    A = T_ACC
    active = [a for a in D.accounts if a["Status"] == "Active"]
    thr = 250000
    r = 4

    def task(r, num, text, hint, height=32):
        ws.write(r, 0, num, bk.f(bold=True, font_color="white", bg_color=C_HDR, align="center"))
        ws.merge_range(r, 1, r, 12, text, bk.f(bold=True, text_wrap=True, valign="top"))
        ws.set_row(r, height)
        ws.merge_range(r + 1, 1, r + 1, 12, "Pista: " + hint, bk.f(italic=True, font_color="#595959",
                                                                   bg_color=C_HINT, text_wrap=True, font_size=9))
        ws.set_row(r + 1, 28)
        return r + 2

    def space(r, rows, cols, c0=1):
        for i in range(rows):
            for j in range(cols):
                if i or j:
                    ws.write_blank(r + i, c0 + j, None, bk.f(bg_color="#F7F7F7"))

    inds = sorted({a["Industry"] for a in D.accounts})
    r = task(r, "7.1", "Lista ordenada (A-Z) de industrias distintas en tbl_Accounts.",
             "SORT(UNIQUE(tbl_Accounts[Industry]))")
    bk.answer(ws, r, 1, f'SORT(UNIQUE({col(A, "Industry")}))')
    space(r, len(inds), 1)
    ws.merge_range(r, 3, r, 5, "7.2 ¿Cuántas industrias distintas hay? →", bk.f(bold=True, align="right"))
    bk.answer(ws, r, 6, f'ROWS(UNIQUE({col(A, "Industry")}))', expected=len(inds), chk=(r, 7))
    ws.write(r + 1, 3, "Pista: ROWS(UNIQUE(...)) o COUNTA(UNIQUE(...))", bk.f(italic=True, font_size=8))
    bk.check_cf(ws, f"H{r + 1}")
    r += len(inds) + 1

    ent = sorted([a for a in active if a["Segment"] == "Enterprise" and a["ARR_USD"] > thr],
                 key=lambda a: -a["ARR_USD"])
    r = task(r, "7.3", f"Todas las cuentas Enterprise ACTIVAS con ARR mayor al umbral de la celda M{r + 3}, "
                       "con las columnas AccountID a ARR_USD, ordenadas por ARR de mayor a menor.",
             "SORT(FILTER(tbl_Accounts[[AccountID]:[ARR_USD]], (tbl_Accounts[Segment]=\"Enterprise\") * "
             "(tbl_Accounts[Status]=\"Active\") * (tbl_Accounts[ARR_USD]>M..)), 8, -1). El * funciona como AND.")
    ws.write(r, 11, "Umbral ARR →", bk.f(bold=True, align="right"))
    ws.write_number(r, 12, thr, bk.f(num_format=NF_USD, bg_color="#DDEBF7", bold=True, border=1))
    thr_ref = f"$M${r + 1}"
    bk.header(ws, r, 1, ["AccountID", "AccountName", "Segment", "Region", "Industry", "Product", "CSM_ID",
                         "ARR_USD"])
    r += 1
    cond = (f'({col(A, "Segment")}="Enterprise")*({col(A, "Status")}="Active")*({col(A, "ARR_USD")}>{thr_ref})')
    bk.answer(ws, r, 1, f'SORT(FILTER(tbl_Accounts[[AccountID]:[ARR_USD]],{cond},"Sin resultados"),8,-1)')
    space(r, max(len(ent), 1), 8)
    for i in range(max(len(ent), 1)):
        ws.write_blank(r + i, 8, None, bk.f(bg_color="#F7F7F7", num_format=NF_USD))
    r += max(len(ent), 1) + 1
    ws.merge_range(r, 1, r, 4, "7.4 ¿Cuántas cuentas cumplen 7.3? (una fórmula) →", bk.f(bold=True, align="right"))
    bk.answer(ws, r, 5, f'IFERROR(ROWS(FILTER({col(A, "AccountID")},{cond})),0)', expected=len(ent), chk=(r, 6))
    ws.write(r, 7, "Pista: IFERROR(ROWS(FILTER(...)), 0) o COUNTIFS (FILTER sin resultados da #CALC!)",
             bk.f(italic=True, font_size=8))
    bk.check_cf(ws, f"G{r + 1}")
    r += 2

    csm0 = D.csms[2]
    names = [a["AccountName"] for a in D.accounts if a["CSM_ID"] == csm0["CSM_ID"] and a["Status"] == "Active"]
    r = task(r, "7.5", "En una sola celda, los nombres de las cuentas activas del CSM indicado, separados por coma "
                       "(útil para un correo o un resumen).",
             "TEXTJOIN(\", \", TRUE, FILTER(tbl_Accounts[AccountName], (tbl_Accounts[CSM_ID]=B..)*(tbl_Accounts[Status]=\"Active\")))")
    ws.write(r, 1, csm0["CSM_ID"], bk.f(bold=True, bg_color="#DDEBF7", border=1))
    ws.write(r, 2, "← ✓ válida solo con este CSM_ID", bk.f(italic=True, font_size=8))
    ws.data_validation(r, 1, r, 1, {"validate": "list", "source": [c["CSM_ID"] for c in D.csms]})
    r += 1
    bk.answer(ws, r, 1, f'TEXTJOIN(", ",TRUE,FILTER({col(A, "AccountName")},({col(A, "CSM_ID")}=B{r})*'
                        f'({col(A, "Status")}="Active")))', expected=", ".join(names), chk=(r, 12))
    ws.set_row(r, 30)
    bk.check_cf(ws, f"M{r + 1}")
    r += 2

    sc = [n["Score"] for n in D.nps]
    nps = (sum(1 for s in sc if s >= 9) - sum(1 for s in sc if s <= 6)) / len(sc) * 100
    r = task(r, "7.6", "Calcula el NPS global en UNA fórmula legible usando LET (variables con nombre).",
             "LET(s, tbl_NPS[Score], p, COUNTIF(s, \">=9\"), d, COUNTIF(s, \"<=6\"), (p − d) / COUNT(s) * 100)")
    bk.answer(ws, r, 1, f'LET(_xlpm.s,{col(T_NPS, "Score")},_xlpm.p,COUNTIF(_xlpm.s,">=9"),_xlpm.d,'
                        f'COUNTIF(_xlpm.s,"<=6"),(_xlpm.p-_xlpm.d)/COUNT(_xlpm.s)*100)', nf=NF_DEC,
              expected=nps, chk=(r, 2), tol=0.05)
    bk.check_cf(ws, f"C{r + 1}")
    r += 2

    acc_x = D.active[len(D.active) // 2]
    r = task(r, "7.7", "Escribe un AccountID en la celda azul y trae con UN solo XLOOKUP las columnas Segment, "
                       "Region, Industry y Product.",
             "XLOOKUP(B.., tbl_Accounts[AccountID], tbl_Accounts[[Segment]:[Product]]) → devuelve 4 columnas")
    ws.write(r, 1, acc_x["AccountID"], bk.f(bold=True, bg_color="#DDEBF7", border=1))
    bk.header(ws, r, 2, ["Segment", "Region", "Industry", "Product"])
    r += 1
    bk.answer(ws, r, 2, bk.XL(f"B{r}", col(A, "AccountID"), "tbl_Accounts[[Segment]:[Product]]")
              if not bk.compat else f'INDEX(tbl_Accounts[[Segment]:[Product]],MATCH(B{r},{col(A, "AccountID")},0),0)')
    space(r, 1, 4, c0=2)
    r += 2

    reds = [h for h in D.health if h["flag"]]
    r = task(r, "7.8", "Desde E3 (usa los nombres HS_ID, HS_Name, HS_ARR, HS_Score, HS_Cat, HS_Days): lista de "
                       "cuentas Red que renuevan en ≤180 días con ID, nombre, ARR y días, ordenada por días a "
                       "renovación (más urgente primero). Depende de que hayas resuelto E3.",
             "SORT(FILTER(HSTACK(HS_ID, HS_Name, HS_ARR, HS_Days), (HS_Cat=\"Red\")*(HS_Days<=180)), 4, 1)")
    bk.header(ws, r, 1, ["AccountID", "AccountName", "ARR", "Días a renovación"])
    r += 1
    bk.answer(ws, r, 1, 'IFERROR(SORT(FILTER(HSTACK(HS_ID,HS_Name,HS_ARR,HS_Days),(HS_Cat="Red")*(HS_Days<=180)),4,1),'
                        '"Sin cuentas")')
    space(r, max(len(reds), 1), 4)
    r += max(len(reds), 1) + 1

    r = task(r, "7.9", "Genera en una fila los 12 meses del análisis (de LastMonth − 11 a LastMonth) con UNA "
                       "fórmula. Úsalo como encabezado de reportes que se actualizan solos.",
             "DATE(YEAR(LastMonth), MONTH(LastMonth) + SEQUENCE(1, 12, -11), 1)  → dale formato mmm yyyy")
    bk.answer(ws, r, 1, "DATE(YEAR(LastMonth),MONTH(LastMonth)+SEQUENCE(1,12,-11),1)", nf=NF_MONTH)
    for j in range(1, 12):
        ws.write_blank(r, 1 + j, None, bk.f(bg_color="#F7F7F7", num_format=NF_MONTH))
    ws.set_column(13, 13, 14)


# ---------------------------------------------------------------------------
# E8 - Tendencias, churn y renovaciones
# ---------------------------------------------------------------------------

def sheet_e8(bk, D):
    ws = bk.ws["E8_Tendencias"]
    bk.title(ws, "E8 · Tendencias de uso, churn y pipeline de renovaciones",
             "Objetivo: construir la vista mensual que liderazgo revisa en el QBR interno: uso total, crecimiento "
             "mes a mes, promedio móvil, cuentas perdidas (churn) y GRR; y el pipeline de renovaciones de los "
             "próximos 6 meses con el ARR en riesgo según tu health score (E3).",
             "Tip: los meses en la columna A son el primer día de cada mes; usa EOMONTH(A,0) como fin de mes en "
             "criterios de fecha: \">=\"&A8 y \"<=\"&EOMONTH(A8,0).", width=19)
    ws.set_column(0, 0, 12)
    ws.set_column(1, 8, 13)
    ws.set_column(9, 9, 16)
    ws.set_column(10, 10, 2)
    ws.set_column(11, 18, 4)
    A, U = T_ACC, T_USE
    labels = ["Month", "Cuentas con uso", "ActiveUsers", "MoM %", "Logins", "Features prom.",
              "Prom. móvil 3m (ActiveUsers)", "Cuentas churned", "ARR churned"]
    hints = ["Dado", "COUNTIFS(tbl_Usage[Month], A8)", "SUMIFS", "C8/C7 − 1 (desde el 2º mes)", "SUMIFS",
             "AVERAGEIFS(...[FeaturesAdopted], ...)", "AVERAGE(C6:C8) desde el 3er mes",
             "COUNTIFS(tbl_Accounts[ChurnDate], \">=\"&A8, ...[ChurnDate], \"<=\"&EOMONTH(A8,0))", "SUMIFS igual"]
    bk.header(ws, 5, 0, labels, hints)
    ws.set_row(6, 62)
    ws.write(5, 9, "Tendencia", bk.f(bold=True, font_color="white", bg_color=C_HDR))
    chk_cols = [1, 2, 3, 4, 5, 6, 7, 8]
    bk.header(ws, 5, 11, [f"✓{xl_col_to_name(c)}" for c in chk_cols])
    first = 7
    tot_au = []
    for i, m in enumerate(D.months):
        r = first + i
        R = r + 1
        us = [u for u in D.usage if u["Month"] == m]
        au = sum(u["ActiveUsers"] for u in us)
        tot_au.append(au)
        ch = [a for a in D.accounts if a["ChurnDate"] and m <= a["ChurnDate"] <= eomonth(m)]
        bk.given(ws, r, 0, m, NF_MONTH)
        bk.answer(ws, r, 1, f'COUNTIFS({col(U, "Month")},$A{R})', expected=len(us), chk=(r, 11))
        bk.answer(ws, r, 2, f'SUMIFS({col(U, "ActiveUsers")},{col(U, "Month")},$A{R})', nf=NF_INT, expected=au,
                  chk=(r, 12))
        if i >= 1:
            bk.answer(ws, r, 3, f"C{R}/C{R - 1}-1", nf=NF_PCT, expected=au / tot_au[i - 1] - 1, chk=(r, 13),
                      tol=0.0005)
        bk.answer(ws, r, 4, f'SUMIFS({col(U, "Logins")},{col(U, "Month")},$A{R})', nf=NF_INT,
                  expected=sum(u["Logins"] for u in us), chk=(r, 14))
        bk.answer(ws, r, 5, f'AVERAGEIFS({col(U, "FeaturesAdopted")},{col(U, "Month")},$A{R})', nf=NF_DEC2,
                  expected=sum(u["FeaturesAdopted"] for u in us) / len(us), chk=(r, 15))
        if i >= 2:
            bk.answer(ws, r, 6, f"AVERAGE(C{R - 2}:C{R})", nf=NF_INT, expected=sum(tot_au[i - 2:i + 1]) / 3,
                      chk=(r, 16), tol=0.5)
        crit = f'{col(A, "ChurnDate")},">="&$A{R},{col(A, "ChurnDate")},"<="&EOMONTH($A{R},0)'
        bk.answer(ws, r, 7, f"COUNTIFS({crit})", expected=len(ch), chk=(r, 17))
        bk.answer(ws, r, 8, f'SUMIFS({col(A, "ARR_USD")},{crit})', nf=NF_USD,
                  expected=sum(a["ARR_USD"] for a in ch), chk=(r, 18))
    last = first + 11
    bk.check_cf(ws, f"L{first + 1}:S{last + 1}")
    ws.add_sparkline(first, 9, {"range": f"E8_Tendencias!C{first + 1}:C{last + 1}", "markers": True})
    ws.add_sparkline(first + 1, 9, {"range": f"E8_Tendencias!E{first + 1}:E{last + 1}", "type": "column"})
    ws.write(first + 2, 9, "↑ ActiveUsers / Logins", bk.f(italic=True, font_size=8))
    ws.conditional_format(f"D{first + 1}:D{last + 1}", {"type": "3_color_scale", "min_color": "#F8696B",
                                                        "mid_color": "#FFFFFF", "max_color": "#63BE7B",
                                                        "mid_type": "num", "mid_value": 0})

    r = last + 2
    churn_arr = sum(a["ARR_USD"] for a in D.accounts if a["ChurnDate"] and D.months[0] <= a["ChurnDate"])
    act_arr = sum(a["ARR_USD"] for a in D.accounts if a["Status"] == "Active")
    bk.given(ws, r, 0, "GRR aprox. 12m", bold=True)
    ws.merge_range(r, 1, r, 4, "= 1 − ARR churned 12m / (ARR activo + ARR churned 12m)",
                   bk.f(italic=True, font_size=8))
    bk.answer(ws, r, 5, f'1-SUM(I{first + 1}:I{last + 1})/(SUMIFS({col(A, "ARR_USD")},{col(A, "Status")},"Active")'
                        f'+SUM(I{first + 1}:I{last + 1}))', nf=NF_PCT,
              expected=1 - churn_arr / (act_arr + churn_arr), chk=(r, 11), tol=0.0005, bold=True)
    bk.check_cf(ws, f"L{r + 1}")

    # renovaciones
    r += 3
    bk.note(ws, r - 1, 0, "Pipeline de renovaciones (próximos 6 meses, cuentas activas)", bold=True,
            font_color=C_TITLE)
    labels = ["Month", "# Renovaciones", "ARR que renueva", "# en Red (E3)", "ARR en Red", "% ARR en riesgo"]
    hints = ["Dado", "COUNTIFS RenewalDate entre A y EOMONTH(A,0), Status Active", "SUMIFS igual",
             "COUNTIFS(HS_Days, \">=\"&(A−AsOfDate), HS_Days, \"<=\"&(EOMONTH(A,0)−AsOfDate), HS_Cat, \"Red\")",
             "SUMIFS(HS_ARR, ...)", "IFERROR(E/C, 0)"]
    bk.header(ws, r, 0, labels, hints)
    ws.set_row(r + 1, 62)
    bk.header(ws, r, 11, ["✓B", "✓C", "✓D", "✓E", "✓F"])
    H = D.health
    for i, m in enumerate(D.future_months):
        rr = r + 2 + i
        R = rr + 1
        ren = [a for a in D.active if m <= a["RenewalDate"] <= eomonth(m)]
        red = [h for h in H if h["cat"] == "Red" and m <= h["acc"]["RenewalDate"] <= eomonth(m)]
        arr_r = sum(a["ARR_USD"] for a in ren)
        arr_red = sum(h["acc"]["ARR_USD"] for h in red)
        bk.given(ws, rr, 0, m, NF_MONTH)
        crit = (f'{col(A, "RenewalDate")},">="&$A{R},{col(A, "RenewalDate")},"<="&EOMONTH($A{R},0),'
                f'{col(A, "Status")},"Active"')
        hcrit = f'HS_Days,">="&($A{R}-AsOfDate),HS_Days,"<="&(EOMONTH($A{R},0)-AsOfDate),HS_Cat,"Red"'
        bk.answer(ws, rr, 1, f"COUNTIFS({crit})", expected=len(ren), chk=(rr, 11))
        bk.answer(ws, rr, 2, f'SUMIFS({col(A, "ARR_USD")},{crit})', nf=NF_USD, expected=arr_r, chk=(rr, 12))
        bk.answer(ws, rr, 3, f"COUNTIFS({hcrit})", expected=len(red), chk=(rr, 13))
        bk.answer(ws, rr, 4, f"SUMIFS(HS_ARR,{hcrit})", nf=NF_USD, expected=arr_red, chk=(rr, 14))
        bk.answer(ws, rr, 5, f"IFERROR(E{R}/C{R},0)", nf=NF_PCT, expected=arr_red / arr_r if arr_r else 0,
                  chk=(rr, 15), tol=0.0005)
    bk.check_cf(ws, f"L{r + 3}:P{r + 8}")
    ws.conditional_format(f"F{r + 3}:F{r + 8}", {"type": "data_bar", "bar_color": "#F8696B",
                                                 "min_type": "num", "min_value": 0,
                                                 "max_type": "num", "max_value": 1})

    chart = bk.wb.add_chart({"type": "line"})
    chart.add_series({"name": "ActiveUsers", "categories": ["E8_Tendencias", first, 0, last, 0],
                      "values": ["E8_Tendencias", first, 2, last, 2], "marker": {"type": "circle"}})
    chart.add_series({"name": "Prom. móvil 3m", "categories": ["E8_Tendencias", first, 0, last, 0],
                      "values": ["E8_Tendencias", first, 6, last, 6], "line": {"dash_type": "dash"}})
    chart.set_title({"name": "Usuarios activos por mes"})
    chart.set_legend({"position": "bottom"})
    chart.set_size({"width": 520, "height": 280})
    ws.insert_chart(r + 10, 0, chart)
    ws.freeze_panes(first, 1)


# ---------------------------------------------------------------------------
# E9 - Dashboard por CSM
# ---------------------------------------------------------------------------

def sheet_e9(bk, D):
    ws = bk.ws["E9_Dashboard_CSM"]
    bk.title(ws, "E9 · Dashboard del CSM (book of business)",
             "Objetivo: un tablero de una página que el CSM abre antes de su 1:1 con su manager. Elige el CSM en "
             "la lista desplegable y todos los KPIs se recalculan. Usa los nombres definidos de E3 (HS_*).",
             "Las ✓ son válidas con el CSM que viene por defecto. Cuando termines, cambia el CSM para ver cómo se "
             "actualiza todo. Bonus: la tabla Top 5 usa LET + FILTER + SORTBY + TAKE.", width=10)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 34)
    ws.set_column(2, 2, 22)
    ws.set_column(3, 3, 4)
    ws.set_column(4, 4, 2)
    ws.set_column(5, 5, 11)
    ws.set_column(6, 6, 24)
    ws.set_column(7, 10, 12)
    csm = D.csms[0]
    ws.write(4, 1, "Selecciona el CSM ▶", bk.f(bold=True, font_size=12, align="right"))
    ws.write(4, 2, csm["CSM_Name"], bk.f(bold=True, font_size=12, bg_color="#DDEBF7", border=2))
    ws.data_validation(4, 2, 4, 2, {"validate": "list", "source": f"=CSMs!$B$2:$B${len(D.csms) + 1}"})
    H = [h for h in D.health if h["acc"]["CSM_ID"] == csm["CSM_ID"]]
    mine = [a for a in D.active if a["CSM_ID"] == csm["CSM_ID"]]
    bk.header(ws, 6, 1, ["KPI", "Valor", "Chk"])
    kpis = [
        ("CSM_ID", bk.XL("$C$5", col(T_CSM, "CSM_Name"), col(T_CSM, "CSM_ID")), csm["CSM_ID"], None),
        ("Team / Manager", f'{bk.XL("$C$5", col(T_CSM, "CSM_Name"), col(T_CSM, "Team"))}&" / "&'
                           f'{bk.XL("$C$5", col(T_CSM, "CSM_Name"), col(T_CSM, "Manager"))}',
         f'{csm["Team"]} / {csm["Manager"]}', None),
        ("# Cuentas activas", f'COUNTIFS({col(T_ACC, "CSM_ID")},$C$8,{col(T_ACC, "Status")},"Active")', len(mine),
         NF_INT),
        ("ARR total", f'SUMIFS({col(T_ACC, "ARR_USD")},{col(T_ACC, "CSM_ID")},$C$8,{col(T_ACC, "Status")},"Active")',
         sum(a["ARR_USD"] for a in mine), NF_USD),
        ("Health score promedio", "IFERROR(AVERAGEIFS(HS_Score,HS_CSM,$C$8),0)",
         sum(h["hs"] for h in H) / len(H) if H else 0, NF_DEC),
        ("# Green", 'COUNTIFS(HS_CSM,$C$8,HS_Cat,"Green")', sum(1 for h in H if h["cat"] == "Green"), NF_INT),
        ("# Yellow", 'COUNTIFS(HS_CSM,$C$8,HS_Cat,"Yellow")', sum(1 for h in H if h["cat"] == "Yellow"), NF_INT),
        ("# Red", 'COUNTIFS(HS_CSM,$C$8,HS_Cat,"Red")', sum(1 for h in H if h["cat"] == "Red"), NF_INT),
        ("Renovaciones en ≤90 días", 'COUNTIFS(HS_CSM,$C$8,HS_Days,"<=90")', sum(1 for h in H if h["days"] <= 90),
         NF_INT),
        ("ARR en riesgo (Red, ≤180 días)", 'SUMIFS(HS_ARR,HS_CSM,$C$8,HS_Cat,"Red",HS_Days,"<=180")',
         sum(h["acc"]["ARR_USD"] for h in H if h["flag"]), NF_USD),
        ("Tickets P1/P2 abiertos", "SUMIFS(HS_Tickets,HS_CSM,$C$8)", sum(h["tk"] for h in H), NF_INT),
    ]
    for i, (lab, fml, exp, nf) in enumerate(kpis):
        r = 7 + i
        ws.write(r, 1, lab, bk.f(bold=True, border=1, border_color="#BFBFBF"))
        bk.answer(ws, r, 2, fml, nf=nf, expected=exp, chk=(r, 3), tol=0.05, bold=True)
    bk.check_cf(ws, f"D8:D{7 + len(kpis)}")
    ws.write(7 + len(kpis), 1, "Pista: C8 se usa en los demás KPIs. Los conteos G/Y/R usan HS_CSM y HS_Cat.",
             bk.f(italic=True, font_size=8))

    ws.merge_range(6, 5, 6, 10, "Top 5 cuentas con peor health score", bk.f(bold=True, font_color="white",
                                                                           bg_color=C_HDR))
    bk.header(ws, 7, 5, ["AccountID", "AccountName", "ARR", "Health", "Categoría", "Días renov."])
    bk.answer(ws, 8, 5, 'IFERROR(LET(_xlpm.f,HS_CSM=$C$8,_xlpm.t,FILTER(HSTACK(HS_ID,HS_Name,HS_ARR,HS_Score,'
                        'HS_Cat,HS_Days),_xlpm.f),TAKE(SORTBY(_xlpm.t,FILTER(HS_Score,_xlpm.f),1),5)),"")')
    for i in range(5):
        for j in range(6):
            if i or j:
                ws.write_blank(8 + i, 5 + j, None, bk.f(bg_color="#F7F7F7",
                                                        num_format=NF_USD if j == 2 else None))
    ws.merge_range(13, 5, 13, 10, "Pista: LET + FILTER(HSTACK(...)) + SORTBY + TAKE → fórmula completa debajo "
                                  "de la gráfica ↓", bk.f(italic=True, font_size=8, bg_color=C_HINT))
    bk.note(ws, 29, 5, "Pista Top 5: LET(f, HS_CSM=$C$8, t, FILTER(HSTACK(HS_ID, HS_Name, HS_ARR, HS_Score, "
                       "HS_Cat, HS_Days), f), TAKE(SORTBY(t, FILTER(HS_Score, f), 1), 5)). Una sola fórmula en F9; "
                       "el resultado se derrama a F9:K13.", merge_to=10, rows=3, italic=True, bg_color=C_HINT,
            font_size=9)
    ws.conditional_format("J9:J13", {"type": "cell", "criteria": "==", "value": '"Red"',
                                     "format": bk.f(bg_color="#FFC7CE", font_color="#9C0006")})
    ws.conditional_format("J9:J13", {"type": "cell", "criteria": "==", "value": '"Yellow"',
                                     "format": bk.f(bg_color="#FFEB9C", font_color="#9C5700")})

    chart = bk.wb.add_chart({"type": "column"})
    chart.add_series({"name": "Cuentas", "categories": ["E9_Dashboard_CSM", 12, 1, 14, 1],
                      "values": ["E9_Dashboard_CSM", 12, 2, 14, 2], "data_labels": {"value": True},
                      "points": [{"fill": {"color": "#70AD47"}}, {"fill": {"color": "#FFC000"}},
                                 {"fill": {"color": "#C00000"}}]})
    chart.set_title({"name": "Distribución de health"})
    chart.set_legend({"none": True})
    chart.set_size({"width": 470, "height": 250})
    ws.insert_chart(15, 5, chart)
    bk.note(ws, 21, 1, "Para la entrevista: ¿qué 3 acciones le sugerirías a este CSM esta semana con base en el "
                       "tablero? ¿Qué dato le falta al tablero para ser más útil (p. ej. fecha del último QBR, "
                       "sponsor ejecutivo, oportunidades de upsell)?", merge_to=2, rows=5, italic=True)


# ---------------------------------------------------------------------------
# E10 - Power Query (texto) y Entrevista
# ---------------------------------------------------------------------------

def sheet_e10(bk, D):
    ws = bk.ws["E10_PowerQuery"]
    bk.title(ws, "E10 · Extra: Power Query (sin verificación automática)",
             "Objetivo: automatizar la ingesta de datos para que el reporte mensual se actualice con un clic en vez "
             "de copiar/pegar. Junto a este archivo hay una carpeta csv/ con las mismas tablas.", width=8)
    ws.set_column(0, 0, 5)
    ws.set_column(1, 1, 110)
    steps = [
        "Datos > Obtener datos > Desde archivo > Desde texto/CSV > csv/Accounts.csv. Revisa que Power Query detecte "
        "bien los tipos (fechas, números). Cierra y carga como Tabla.",
        "Repite con Usage_Monthly.csv. En el editor: Inicio > Combinar consultas (Merge) con Accounts por AccountID "
        "(Left Outer) y expande Segment, Region y CSM_ID.",
        "Agrupa (Transformar > Agrupar por) Month + Segment → suma de ActiveUsers. Compara con tu pivot P2 de E6.",
        "Limpieza en Power Query: carga el rango del ejercicio E5 (columnas A:F) → Formato > Recortar, Poner en "
        "mayúsculas, Reemplazar valores ('$', 'USD'), Cambiar tipo usando configuración regional, Quitar duplicados. "
        "Compara el resultado con tus fórmulas de E5.",
        "Report_Usage_Log.csv: agrega columna personalizada Days_Since = Duration.Days(#date(...) - [Date]) y agrupa "
        "por ReportID con Recuento, Recuento distinto de UserEmail y Máximo de Date (= auditoría E4 sin fórmulas).",
        "Desde carpeta: pon varios CSV mensuales en una carpeta y usa Obtener datos > Desde carpeta > Combinar. Es "
        "el patrón típico para reportes contractuales mensuales.",
        "Ctrl+Alt+F5 para actualizar todo. Documenta los pasos aplicados (panel derecho): es tu 'audit trail'.",
    ]
    for i, s in enumerate(steps):
        ws.write(4 + i, 0, i + 1, bk.f(bold=True, align="center", valign="top"))
        bk.note(ws, 4 + i, 1, s, height=34)


INTERVIEW = [
    ("Excel técnico", [
        "Walk me through how you would build a customer health score. Which inputs and weights would you choose, "
        "and how would you validate it?",
        "XLOOKUP vs VLOOKUP vs INDEX/MATCH: when would you use each one?",
        "How do you find and handle duplicates in a 10,000-row export from Salesforce?",
        "A pivot table total doesn't match the number in Salesforce. How do you troubleshoot it?",
        "How would you calculate unique users per report? (UNIQUE/FILTER, pivot with Distinct Count via Data Model, "
        "Power Query)",
        "What's the difference between SUMIFS and SUMPRODUCT? Give an example where you need SUMPRODUCT.",
    ]),
    ("Auditoría de reportes y datos", [
        "We have ~40 recurring reports. How would you decide which ones to retire, consolidate or automate?",
        "What usage metrics would you track for internal reports (views, unique users, recency, exports, time spent, "
        "cost to maintain)?",
        "A contractual report shows zero views. What do you do before recommending any change?",
        "How do you make sure the data you deliver is accurate? Describe your QA checklist.",
    ]),
    ("Trabajo con CSMs y liderazgo", [
        "How would you partner with CSMs to understand which reports actually help them manage account health?",
        "A CSM disagrees with the health score of one of their accounts (it shows Red but they say it is fine). "
        "How do you handle it?",
        "Tell me about a time you turned data into a clear story for a non-technical audience.",
        "You receive five urgent ad hoc requests on the same day. How do you prioritize?",
        "What would your first 30/60/90 days look like in this role?",
    ]),
    ("Preguntas para hacerles a ellos", [
        "Which tools are the source of truth today (Salesforce, Gainsight, product database, BI tool)?",
        "How is the current health score defined and how much do CSMs trust it?",
        "Which reports take the most time to produce today?",
        "How will success be measured for this role in the first 6 months?",
    ]),
]

GUIDE = ("Guía (versión soluciones): estructura STAR para conductuales; para preguntas técnicas explica el porqué, "
         "menciona validaciones (totales de control, comparación con la fuente, muestreo de filas) y cierra con el "
         "impacto de negocio (horas ahorradas, ARR protegido, mejores decisiones de los CSMs).")


def sheet_interview(bk, D):
    ws = bk.ws["Entrevista"]
    bk.title(ws, "Preguntas de práctica para la entrevista",
             "Practica en voz alta y en inglés. Usa los resultados de tus ejercicios (E3, E4, E5) como ejemplos "
             "concretos con números.", width=3)
    ws.set_column(0, 0, 4)
    ws.set_column(1, 1, 110)
    r = 4
    for section, qs in INTERVIEW:
        ws.merge_range(r, 0, r, 1, section, bk.f(bold=True, font_color="white", bg_color=C_HDR))
        r += 1
        for i, q in enumerate(qs):
            ws.write(r, 0, i + 1, bk.f(align="center", valign="top"))
            bk.note(ws, r, 1, q, height=30)
            r += 1
        r += 1
    if bk.solved:
        bk.note(ws, r, 1, GUIDE, height=45, italic=True, bg_color=C_SOL)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def build(path, D, solved, compat=False):
    bk = Book(path, solved, compat)
    write_data_sheets(bk, D)
    sheet_leeme(bk, D)
    sheet_dictionary(bk, D)
    sheet_e1(bk, D)
    sheet_e2(bk, D)
    sheet_e3(bk, D)
    sheet_e4(bk, D)
    sheet_e5(bk, D)
    sheet_e6(bk, D)
    sheet_e7(bk, D)
    sheet_e8(bk, D)
    sheet_e9(bk, D)
    sheet_e10(bk, D)
    sheet_interview(bk, D)
    finalize_progress(bk)
    for name in EXERCISES:
        bk.ws[name].set_tab_color("#FFC000" if not solved else "#70AD47")
    bk.close()
    return sum(bk.n_checks.values())


def last_day_prev_month(today):
    return today.replace(day=1) - dt.timedelta(days=1)


def main():
    ap = argparse.ArgumentParser(description="Genera libros de Excel para practicar CS Ops / análisis de datos.")
    ap.add_argument("--salida", default="salida", help="Carpeta de salida (default: salida)")
    ap.add_argument("--semilla", type=int, default=42, help="Semilla aleatoria; cambia para obtener otros datos")
    ap.add_argument("--cuentas", type=int, default=150, help="Número de cuentas (min 60, default 150)")
    ap.add_argument("--fecha-corte", default=None,
                    help="Fecha de corte AAAA-MM-DD (default: último día del mes anterior)")
    ap.add_argument("--compat-libreoffice", action="store_true", help=argparse.SUPPRESS)
    args = ap.parse_args()

    as_of = (dt.date.fromisoformat(args.fecha_corte) if args.fecha_corte
             else last_day_prev_month(dt.date.today()))
    n = max(60, args.cuentas)
    D = generate(args.semilla, n, as_of)
    D.seed = args.semilla
    compute_health(D)

    os.makedirs(args.salida, exist_ok=True)
    p1 = os.path.join(args.salida, "CS_Ops_Practica.xlsx")
    p2 = os.path.join(args.salida, "CS_Ops_Soluciones.xlsx")
    n_checks = build(p1, D, solved=False, compat=args.compat_libreoffice)
    build(p2, D, solved=True, compat=args.compat_libreoffice)
    write_csvs(D, os.path.join(args.salida, "csv"))

    print(f"Fecha de corte: {as_of}  |  semilla: {args.semilla}  |  cuentas: {n}")
    print(f"  {p1}  ({n_checks} verificaciones automáticas)")
    print(f"  {p2}")
    print(f"  {os.path.join(args.salida, 'csv')}/  (tablas en CSV para Power Query)")


if __name__ == "__main__":
    main()
