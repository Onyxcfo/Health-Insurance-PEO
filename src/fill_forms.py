# -*- coding: utf-8 -*-
"""Fill the three NY DOH HCRA forms for Onyx Accounting Group LLC."""
import pymupdf, shutil, os

U   = '/root/.claude/uploads/8e61193d-64dc-57a8-9fa0-d6acf524ead7/'
OUT = 'angle_forms/'

ENTITY  = "Onyx Accounting Group LLC"
FEIN    = "85-3909741"
ADDR1   = "14614 N Kierland Blvd Suite N220"
ADDR2   = "Scottsdale, AZ 85254"
CITY, STATE, ZIP = "Scottsdale", "AZ", "85254"
PHONE   = "480-442-3119"
CONTACT = "Josephine Mack"
EMAIL   = "jmack@onyxcfo.com"
EFF     = "10/01/2026"
TPA     = "Adrem Administrators"
TPAFEIN = "87-2258484"
SIGNER  = "Steven Nikolov"
TITLE   = "Principal"
FONT, SZ, COL = "helv", 10.5, (0, 0, 0.6)

def put(page, x, y, text, size=SZ, rotate=0):
    page.insert_text((x, y), text, fontname=FONT, fontsize=size,
                     color=COL, rotate=rotate)

# ===================== DOH-4399 =====================
d = pymupdf.open(U + '4b3cf1ed-Angle_DOH-4399.pdf')

p = d[2]                                    # form page 1 of 5
put(p, 220, 144.0, EFF)
put(p, 220, 192.3, FEIN)
put(p, 220, 216.7, ENTITY)
put(p, 220, 241.0, "N/A")
put(p, 220, 265.4, ADDR1)
put(p, 220, 289.8, ADDR2)
put(p, 220, 314.2, CONTACT)
put(p, 220, 338.6, PHONE)
put(p, 220, 362.9, EMAIL)
# TPA name / FEIN already pre-printed by Angle

p = d[3]                                    # form page 2 of 5 - Title only; Steven signs
r = p.search_for("Title")
tx = r[0].x1 + 4 if r else 436
put(p, tx, 508.0, TITLE)

p = d[4]                                    # form page 3 of 5 - Coverage Information (rotated 90)
put(p, 101.5, 615.0, ENTITY,  rotate=90)
put(p, 101.5, 305.0, FEIN,    rotate=90)
put(p, 122.2, 610.0, TPA,     rotate=90)
put(p, 122.2, 262.0, TPAFEIN, rotate=90)
# Row 5 (Self-Insured Fund WITH a TPA) x Self-Insured Coverage column
put(p, 393.5, 387.0, "X", size=14, rotate=90)

d.save(OUT + 'DOH-4399_Payor_Election_Onyx.pdf'); d.close()

# ===================== DOH-4264 =====================
d = pymupdf.open(U + '5094bce9-Angle_DOH-4264.pdf')
p = d[1]
put(p, 136.9, 124.2, "X", size=10)          # New Request
put(p,  40.0, 185.5, ENTITY)
put(p, 302.0, 210.8, FEIN)
put(p,  55.7, 275.2, "X", size=10)          # Public Goods Pool
put(p, 155.0, 415.5, SIGNER)                # Name (please print)
put(p,  72.0, 441.0, TITLE)
put(p, 125.0, 466.3, PHONE)
put(p,  90.0, 491.5, ADDR1)
put(p,  70.0, 542.0, CITY)
put(p, 300.0, 542.0, STATE)
put(p, 428.0, 542.0, ZIP)
put(p, 125.0, 567.5, EMAIL)
# Signature and Date left blank for Steven
d.save(OUT + 'DOH-4264_Electronic_Filing_ID_Onyx.pdf'); d.close()

# ===================== DOH-4403 =====================
d = pymupdf.open(U + 'ef5a506a-Angle_DOH-4403.pdf')
p = d[1]
put(p, 252.0, 139.8, EFF)
put(p, 110.0, 187.8, ENTITY)
put(p, 470.0, 187.8, FEIN)
put(p, 120.0, 210.2, CONTACT)
put(p, 470.0, 210.2, PHONE)
put(p, 125.0, 325.5, "N/A - initial HCRA election, no prior TPA", size=9.5)
put(p, 470.0, 325.5, "N/A", size=9.5)
# Section II pre-printed by Angle; Section III checkboxes left exactly as Angle set them
d.save(OUT + 'DOH-4403_TPA_Status_Change_Onyx.pdf'); d.close()

print("written:", sorted(os.listdir(OUT)))
