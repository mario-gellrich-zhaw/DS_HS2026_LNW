"""
Generate cinema.db and cinema_data.xlsx for the DS HS2026 exam.

WARNING: Do not run this script on the exam data. The current cinema.db and
cinema_data.xlsx in DS_HS2026_LNW_I_Examples/AP01 were extended after
generation (cinema C006 «Kino Aurora» without screenings, customers
CU0501–CU0505 with inconsistent membership_type, sheets «Description» and
«ER-Diagram»). This script does not reproduce them and would overwrite them.

Schema:
  movies      (movieid, title, genre, duration_min, release_year, director, age_rating)
  cinemas     (cinemaid, name, city, total_seats)
  screenings  (screeningid, movieid, cinemaid, screening_datetime, base_price)
  customers   (customerid, firstname, lastname, email, birthdate, membership_type)
  tickets     (ticketid, screeningid, customerid, purchase_datetime, seat_number, paid_price)
  staff       (staffid, firstname, lastname, cinemaid, role, hire_date)
"""

import sqlite3, random, os
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

random.seed(42)

# ── helpers ──────────────────────────────────────────────────────────────────
def rand_date(start, end):
    delta = end - start
    return start + timedelta(days=random.randint(0, delta.days))

def rand_datetime(start, end):
    delta = end - start
    secs  = int(delta.total_seconds())
    return start + timedelta(seconds=random.randint(0, secs))

# ── raw data ──────────────────────────────────────────────────────────────────
MOVIES = [
    ("M001","Inception",           "Sci-Fi",  148, 2010, "Christopher Nolan",  12),
    ("M002","The Dark Knight",     "Action",  152, 2008, "Christopher Nolan",  12),
    ("M003","Interstellar",        "Sci-Fi",  169, 2014, "Christopher Nolan",  12),
    ("M004","Parasite",            "Thriller",132, 2019, "Bong Joon-ho",       16),
    ("M005","Everything Everywhere","Comedy", 139, 2022, "Daniel Kwan",        12),
    ("M006","Oppenheimer",         "Drama",   180, 2023, "Christopher Nolan",  12),
    ("M007","Barbie",              "Comedy",  114, 2023, "Greta Gerwig",        6),
    ("M008","The Godfather",       "Drama",   175, 1972, "Francis Coppola",    16),
    ("M009","Pulp Fiction",        "Crime",   154, 1994, "Quentin Tarantino",  18),
    ("M010","The Matrix",          "Sci-Fi",  136, 1999, "Wachowski Sisters",  16),
    ("M011","Spirited Away",       "Animation",125,2001, "Hayao Miyazaki",      6),
    ("M012","Dune: Part Two",      "Sci-Fi",  166, 2024, "Denis Villeneuve",   12),
    ("M013","Poor Things",         "Drama",   141, 2023, "Yorgos Lanthimos",   18),
    ("M014","Past Lives",          "Romance", 106, 2023, "Celine Song",         6),
    ("M015","Joker",               "Thriller",122, 2019, "Todd Phillips",      16),
]

CINEMAS = [
    ("C001","Arthouse Metropol", "Zürich",  180),
    ("C002","Pathé Küchlin",     "Basel",   450),
    ("C003","Cineplexx",         "Bern",    320),
    ("C004","Rex Kino",          "Luzern",  140),
    ("C005","Kinok",             "St. Gallen",110),
]

STAFF_ROLES = ["Kassier:in","Vorführer:in","Reinigung","Manager:in","Techniker:in"]
FIRST_NAMES = ["Lena","Noah","Emma","Liam","Mia","Felix","Anna","Jan","Sara","Tom",
               "Julia","Paul","Laura","Max","Nina","David","Lea","Simon","Marie","Lucas"]
LAST_NAMES  = ["Müller","Schmidt","Fischer","Weber","Meier","Keller","Huber","Wolf",
               "Zimmermann","Braun","Hartmann","Lange","Krüger","Richter","Bauer"]
MEMBERSHIP  = ["Standard","Premium","Student","Senior"]
GENRES      = list({m[2] for m in MOVIES})

def gen_staff(n=30):
    rows = []
    for i in range(1, n+1):
        sid    = f"S{i:03d}"
        fn     = random.choice(FIRST_NAMES)
        ln     = random.choice(LAST_NAMES)
        cinema = random.choice(CINEMAS)[0]
        role   = random.choice(STAFF_ROLES)
        hired  = rand_date(datetime(2015,1,1), datetime(2024,6,1))
        rows.append((sid, fn, ln, cinema, role, hired.strftime("%Y-%m-%d")))
    return rows

def gen_customers(n=500):
    rows = []
    used_emails = set()
    for i in range(1, n+1):
        cid  = f"CU{i:04d}"
        fn   = random.choice(FIRST_NAMES)
        ln   = random.choice(LAST_NAMES)
        base = f"{fn.lower()}.{ln.lower()}"
        email = f"{base}{random.randint(1,999)}@example.com"
        while email in used_emails:
            email = f"{base}{random.randint(1,9999)}@example.com"
        used_emails.add(email)
        bd   = rand_date(datetime(1960,1,1), datetime(2006,1,1))
        mem  = random.choice(MEMBERSHIP)
        rows.append((cid, fn, ln, email, bd.strftime("%Y-%m-%d"), mem))
    return rows

def gen_screenings(n=300):
    rows = []
    used = set()
    for i in range(1, n+1):
        sid  = f"SC{i:04d}"
        mid  = random.choice(MOVIES)[0]
        cid  = random.choice(CINEMAS)[0]
        key  = (mid, cid)
        # allow repeats but vary datetime
        dt   = rand_datetime(datetime(2024,1,1), datetime(2025,6,30,23,0))
        # round to next half hour
        dt   = dt.replace(minute=30*(dt.minute//30), second=0, microsecond=0)
        price= round(random.choice([12.0, 14.0, 14.5, 16.0, 18.0, 20.0, 22.0]), 2)
        rows.append((sid, mid, cid, dt.strftime("%Y-%m-%d %H:%M:%S"), price))
    return rows

def gen_tickets(screenings, customers, n=2000):
    rows = []
    used_seats = {}  # (screeningid, seat) -> ticketid
    sc_ids  = [s[0] for s in screenings]
    cu_ids  = [c[0] for c in customers]
    for i in range(1, n+1):
        tid  = f"T{i:05d}"
        scid = random.choice(sc_ids)
        cuid = random.choice(cu_ids)
        seat = f"{random.choice('ABCDEFGHIJ')}{random.randint(1,20)}"
        # avoid duplicate seats per screening
        attempts = 0
        while (scid, seat) in used_seats and attempts < 10:
            seat = f"{random.choice('ABCDEFGHIJ')}{random.randint(1,20)}"
            attempts += 1
        if (scid, seat) in used_seats:
            continue
        used_seats[(scid, seat)] = tid
        # find base price for screening
        sc_price = next((s[4] for s in screenings if s[0] == scid), 14.0)
        # members get discount
        paid = round(sc_price * random.choice([1.0, 1.0, 0.85, 0.9]), 2)
        pdt  = rand_datetime(datetime(2024,1,1), datetime(2025,6,30))
        rows.append((tid, scid, cuid, pdt.strftime("%Y-%m-%d %H:%M:%S"), seat, paid))
    return rows

# ── build ─────────────────────────────────────────────────────────────────────
staff      = gen_staff(30)
customers  = gen_customers(500)
screenings = gen_screenings(300)
tickets    = gen_tickets(screenings, customers, 2000)

# Output goes to the AP01 work package folder (DS_HS2026_LNW_I_Examples/AP01)
OUT_DIR = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                        "..", "..", "DS_HS2026_LNW_I_Examples", "AP01"))
DB_PATH = os.path.join(OUT_DIR, "cinema.db")
XL_PATH = os.path.join(OUT_DIR, "cinema_data.xlsx")

# ── SQLite ────────────────────────────────────────────────────────────────────
if os.path.exists(DB_PATH):
    os.remove(DB_PATH)

conn = sqlite3.connect(DB_PATH)
cur  = conn.cursor()

cur.executescript("""
CREATE TABLE movies (
    movieid       TEXT PRIMARY KEY,
    title         TEXT NOT NULL,
    genre         TEXT NOT NULL,
    duration_min  INTEGER,
    release_year  INTEGER,
    director      TEXT,
    age_rating    INTEGER
);

CREATE TABLE cinemas (
    cinemaid    TEXT PRIMARY KEY,
    name        TEXT NOT NULL,
    city        TEXT NOT NULL,
    total_seats INTEGER
);

CREATE TABLE staff (
    staffid    TEXT PRIMARY KEY,
    firstname  TEXT NOT NULL,
    lastname   TEXT NOT NULL,
    cinemaid   TEXT,
    role       TEXT,
    hire_date  DATE,
    FOREIGN KEY (cinemaid) REFERENCES cinemas(cinemaid)
);

CREATE TABLE screenings (
    screeningid        TEXT PRIMARY KEY,
    movieid            TEXT NOT NULL,
    cinemaid           TEXT NOT NULL,
    screening_datetime DATETIME NOT NULL,
    base_price         REAL,
    FOREIGN KEY (movieid)  REFERENCES movies(movieid),
    FOREIGN KEY (cinemaid) REFERENCES cinemas(cinemaid)
);

CREATE TABLE customers (
    customerid      TEXT PRIMARY KEY,
    firstname       TEXT NOT NULL,
    lastname        TEXT NOT NULL,
    email           TEXT UNIQUE,
    birthdate       DATE,
    membership_type TEXT
);

CREATE TABLE tickets (
    ticketid          TEXT PRIMARY KEY,
    screeningid       TEXT NOT NULL,
    customerid        TEXT NOT NULL,
    purchase_datetime DATETIME,
    seat_number       TEXT,
    paid_price        REAL,
    FOREIGN KEY (screeningid) REFERENCES screenings(screeningid),
    FOREIGN KEY (customerid)  REFERENCES customers(customerid)
);
""")

cur.executemany("INSERT INTO movies VALUES (?,?,?,?,?,?,?)", MOVIES)
cur.executemany("INSERT INTO cinemas VALUES (?,?,?,?)", CINEMAS)
cur.executemany("INSERT INTO staff VALUES (?,?,?,?,?,?)", staff)
cur.executemany("INSERT INTO screenings VALUES (?,?,?,?,?)", screenings)
cur.executemany("INSERT INTO customers VALUES (?,?,?,?,?,?)", customers)
cur.executemany("INSERT INTO tickets VALUES (?,?,?,?,?,?)", tickets)

conn.commit()

# verify
for tbl in ["movies","cinemas","staff","screenings","customers","tickets"]:
    n = cur.execute(f"SELECT COUNT(*) FROM {tbl}").fetchone()[0]
    print(f"  {tbl:<12}: {n:>5} rows")

conn.close()
print(f"\nSQLite → {DB_PATH}")

# ── Excel ─────────────────────────────────────────────────────────────────────
wb = openpyxl.Workbook()
wb.remove(wb.active)  # remove default sheet

HEADER_FILL  = PatternFill("solid", fgColor="2E4057")
HEADER_FONT  = Font(color="FFFFFF", bold=True)
ALT_FILL     = PatternFill("solid", fgColor="EAF0FB")

tables = {
    "movies":     (["movieid","title","genre","duration_min","release_year","director","age_rating"], MOVIES),
    "cinemas":    (["cinemaid","name","city","total_seats"], CINEMAS),
    "staff":      (["staffid","firstname","lastname","cinemaid","role","hire_date"], staff),
    "screenings": (["screeningid","movieid","cinemaid","screening_datetime","base_price"], screenings),
    "customers":  (["customerid","firstname","lastname","email","birthdate","membership_type"], customers),
    "tickets":    (["ticketid","screeningid","customerid","purchase_datetime","seat_number","paid_price"], tickets),
}

for sheet_name, (headers, rows) in tables.items():
    ws = wb.create_sheet(title=sheet_name)
    # header row
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font  = HEADER_FONT
        cell.fill  = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    # data rows
    for r_idx, row in enumerate(rows, 2):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            if r_idx % 2 == 0:
                cell.fill = ALT_FILL
    # auto-width
    for col in ws.columns:
        max_len = max(len(str(c.value)) if c.value else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

wb.save(XL_PATH)
print(f"Excel  → {XL_PATH}")
