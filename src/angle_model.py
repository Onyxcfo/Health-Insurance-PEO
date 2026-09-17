# -*- coding: utf-8 -*-
"""Onyx Accounting Group - Angle Health + Questco cost model."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

F="Arial"
INK=Font(name=F,size=10); BLD=Font(name=F,size=10,bold=True)
SM=Font(name=F,size=9,color="404040"); SUB=Font(name=F,size=9,italic=True,color="595959")
BLU=Font(name=F,size=10,color="0000FF"); BLUB=Font(name=F,size=10,bold=True,color="0000FF")
GRN=Font(name=F,size=10,color="008000"); TTL=Font(name=F,size=15,bold=True)
HDF=Font(name=F,size=9,bold=True,color="FFFFFF"); SECF=Font(name=F,size=11,bold=True,color="1D5B72")
HFILL=PatternFill("solid",fgColor="1D5B72"); YEL=PatternFill("solid",fgColor="FFFF00")
BAND=PatternFill("solid",fgColor="EFF3F5"); TOTFILL=PatternFill("solid",fgColor="E4E9EA")
thin=Side(style="thin",color="BFBFBF"); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
BOT=Border(bottom=Side(style="thin",color="1D5B72"))
CUR='$#,##0;($#,##0);-'; CUR2='$#,##0.00;($#,##0.00);-'; PCT='0.0%'
WRAP=Alignment(wrap_text=True,vertical="top"); CTR=Alignment(horizontal="center")

# vendor, plan, hsa, ded_i, ded_f, oop_i, oop_f, plan_pays, pcp, spec, rx, er, EE, ES, EC, FAM, note
ANGLE=[
 ("ANG HDHP 3400/5000","Y",3400,6800,5000,10000,0.80,"20% after ded","20% after ded","20% after ded (all tiers)","20% after ded",
  439.03,921.97,834.16,1361.00,"CapFi's recommended base plan. Sets the minimum-funding floor."),
 ("ANG TRAD 2000/4000","N",2000,4000,4000,8000,0.80,"$20","$50","$20 / $60","$250",
  492.11,1033.43,934.01,1525.54,"Tier rates DERIVED from the ratios in the two documented plans. Confirm with CapFi before quoting to employees."),
 ("ANG TRAD 1000/2000","N",1000,2000,2000,4000,0.80,"$10","$30","$10 / $30","$200",
  546.33,1147.29,1038.02,1693.61,"Best plan design in the whole field - $1,000 deductible, $2,000 out-of-pocket maximum."),
]
TIERS=[("Employee only","EE",300.0),("Employee + spouse","ES",600.0),
       ("Employee + child(ren)","EC",600.0),("Family","FAM",1000.0)]
EMP=[("Steven Nikolov",186000,"Family","FAM","Yes","100% owner. >2% S-corp shareholder - premiums cannot run pretax through the cafeteria plan (IRC 1372)."),
     ("Lisa Danforth",178094,"Employee only","EE","Yes",""),
     ("Jessica Bumphus",135000,"Family","FAM","Yes","New hire."),
     ("Christine Johnson",79228,"Employee + spouse","ES","Yes","Lowest salary, second-highest tier - feels the deduction hardest."),
     ("Josephine Mack",80000,"Family","FAM","No","Waiving. Shown so the model can be re-run with her enrolled - change to Yes.")]

wb=openpyxl.Workbook()
def sec(ws,row,text,span=8):
    c=ws.cell(row=row,column=1,value=text); c.font=SECF
    for col in range(1,span+1): ws.cell(row=row,column=col).border=BOT
def hdr(ws,row,labels,start=1,h=30):
    for i,l in enumerate(labels):
        c=ws.cell(row=row,column=start+i,value=l); c.font=HDF; c.fill=HFILL; c.border=BOX
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.row_dimensions[row].height=h

# ================= 1. CONTRIBUTIONS (the levers) =================
a=wb.active; a.title="Contributions"
a["A1"]="Onyx Accounting Group, LLC - Angle Health + Questco"; a["A1"].font=TTL
a["A2"]="Medical placed with Angle Health through CapFi. All other PEO services through Questco. Medical effective 10/01/2026."
a["A2"].font=SUB
a["A3"]="EDIT THE YELLOW CELLS ONLY. Every other number on every tab recalculates from them."
a["A3"].font=Font(name=F,size=10,bold=True,color="8E2F2A")

sec(a,5,"1.  ONYX MONTHLY CONTRIBUTION BY TIER   <- this is the lever")
hdr(a,6,["Tier","Code","Onyx pays $/mo","Onyx pays $/yr","Angle minimum funding $/mo","Clears the minimum?"])
for i,(name,code,amt) in enumerate(TIERS):
    r=7+i
    a.cell(row=r,column=1,value=name).font=INK
    a.cell(row=r,column=2,value=code).font=INK; a.cell(row=r,column=2).alignment=CTR
    c=a.cell(row=r,column=3,value=amt); c.font=BLUB; c.fill=YEL; c.number_format=CUR2
    c=a.cell(row=r,column=4,value=f"=C{r}*12"); c.font=INK; c.number_format=CUR
    c=a.cell(row=r,column=5,value="=$C$16"); c.font=GRN; c.number_format=CUR2
    c=a.cell(row=r,column=6,value=f'=IF(C{r}>=E{r},"Yes","NO - below Angle minimum")')
    c.font=INK; c.alignment=CTR
    for col in range(1,7): a.cell(row=r,column=col).border=BOX
a["A12"]="Fixed dollars, medical only. Dental and vision are not offered by Angle; life, LTD and STD are quoted separately by Questco on the Questco tab."
a["A12"].font=SUB

sec(a,14,"2.  ANGLE MINIMUM FUNDING REQUIREMENT")
a["A15"]="Percentage of the lowest employee-only rate"; a["A15"].font=INK
c=a["C15"]; c.value=0.50; c.font=BLUB; c.fill=YEL; c.number_format=PCT
a["A16"]="Minimum Onyx must fund, per enrolled employee, per month"; a["A16"].font=BLD
c=a["C16"]; c.value="=C15*MIN('Angle Plans'!L5:L7)"; c.font=BLD; c.number_format=CUR2
a["A17"]="Minimum total funding at current enrollment, per year"; a["A17"].font=INK
c=a["C17"]; c.value="=C16*C34*12"; c.font=INK; c.number_format=CUR
a["A18"]=("Angle's own presentation states minimum funding of $10,536.96/yr - 50% of the HDHP 3400/5000's $439.03 employee-only "
          "rate across four enrolled employees. The formula above lands at $10,536.72; the 24-cent difference is Angle rounding the "
          "monthly floor to $219.52.")
a["A18"].font=SUB
for r in (15,16,17):
    for col in range(1,4): a.cell(row=r,column=col).border=BOX

sec(a,20,"3.  WHO IS ENROLLED   <- change Yes/No to re-run the whole model")
hdr(a,21,["Employee","Annual pay","Medical tier","Code","Enrolling?","Note"])
for i,(n,pay,tier,code,yn,note) in enumerate(EMP):
    r=22+i
    a.cell(row=r,column=1,value=n).font=INK
    c=a.cell(row=r,column=2,value=pay); c.font=BLU; c.number_format=CUR
    a.cell(row=r,column=3,value=tier).font=INK
    a.cell(row=r,column=4,value=code).font=BLU; a.cell(row=r,column=4).alignment=CTR
    c=a.cell(row=r,column=5,value=yn); c.font=BLUB; c.fill=YEL; c.alignment=CTR
    a.cell(row=r,column=6,value=note).font=SM; a.cell(row=r,column=6).alignment=WRAP
    a.row_dimensions[r].height=26
    for col in range(1,7): a.cell(row=r,column=col).border=BOX
dv=DataValidation(type="list",formula1='"Yes,No"',allow_blank=False,showDropDown=False)
a.add_data_validation(dv); dv.add("E22:E26")
EMP_FIRST,EMP_LAST=22,26

sec(a,28,"4.  ENROLLMENT AND COST AT THE CURRENT SETTINGS")
hdr(a,29,["Tier","Code","Enrolled","Onyx $/mo","Onyx $/yr"])
for i,(name,code,_a) in enumerate(TIERS):
    r=30+i
    a.cell(row=r,column=1,value=name).font=INK
    a.cell(row=r,column=2,value=code).font=INK; a.cell(row=r,column=2).alignment=CTR
    c=a.cell(row=r,column=3,value=f'=COUNTIFS($D${EMP_FIRST}:$D${EMP_LAST},B{r},$E${EMP_FIRST}:$E${EMP_LAST},"Yes")')
    c.font=INK; c.alignment=CTR
    c=a.cell(row=r,column=4,value=f"=C{r}*INDEX($C$7:$C$10,MATCH(B{r},$B$7:$B$10,0))"); c.font=INK; c.number_format=CUR2
    c=a.cell(row=r,column=5,value=f"=D{r}*12"); c.font=INK; c.number_format=CUR
    for col in range(1,6): a.cell(row=r,column=col).border=BOX
r=34
a.cell(row=r,column=1,value="TOTAL").font=BLD
c=a.cell(row=r,column=3,value="=SUM(C30:C33)"); c.font=BLD; c.alignment=CTR
c=a.cell(row=r,column=4,value="=SUM(D30:D33)"); c.font=BLD; c.number_format=CUR2
c=a.cell(row=r,column=5,value="=SUM(E30:E33)"); c.font=BLD; c.number_format=CUR
for col in range(1,6):
    a.cell(row=r,column=col).border=BOX; a.cell(row=r,column=col).fill=TOTFILL
for col,w in [("A",42),("B",11),("C",17),("D",17),("E",22),("F",62)]:
    a.column_dimensions[col].width=w
a.sheet_view.showGridLines=False
print("contributions ok")

# ================= 2. ANGLE PLANS =================
p=wb.create_sheet("Angle Plans")
p["A1"]="Angle Health - all three plans"; p["A1"].font=TTL
p["A2"]="Level funded on the Cigna PPO network. Out-of-network is covered at 50% after a separate out-of-network deductible. Coinsurance is the share the PLAN pays after the deductible."
p["A2"].font=SUB
hdr(p,4,["Plan","HSA","Ded individual","Ded family","OOP max individual","OOP max family","Coins (plan pays)",
         "PCP","Specialist","Rx","Emergency room","EE","ES","EC","FAM","Note"])
for i,row in enumerate(ANGLE):
    r=5+i
    name,hsa,di,df,oi,of,cp,pcp,spec,rx,er,ee,es,ec,fam,note=row
    for j,v in enumerate([name,hsa,di,df,oi,of,cp,pcp,spec,rx,er,ee,es,ec,fam,note]):
        c=p.cell(row=r,column=1+j,value=v); c.font=BLU; c.border=BOX
        if j in (2,3,4,5): c.number_format=CUR
        if j in (11,12,13,14): c.number_format=CUR2
        if j==6: c.number_format="0%"
        if j in (1,6): c.alignment=CTR
        if j==15: c.font=SM; c.alignment=WRAP
    p.row_dimensions[r].height=34
    if i%2:
        for j in range(16): p.cell(row=r,column=1+j).fill=BAND
p["A9"]="Every cell above is transcribed from the CapFi / Angle Health presentation, quote 400798, underwritten 7/30/2026."
p["A9"].font=SUB
for col,w in [("A",24),("B",6),("C",13),("D",13),("E",15),("F",15),("G",12),("H",15),("I",15),
              ("J",22),("K",15),("L",11),("M",11),("N",11),("O",11),("P",50)]:
    p.column_dimensions[col].width=w
p.freeze_panes="B5"; p.sheet_view.showGridLines=False
print("plans ok")
wb.save("_stage1.xlsx"); print("stage 1 saved")

# ================= 3. COST BY TIER =================
PL="'Angle Plans'!"
PN=f"{PL}$A$5:$A$7"; PPREM=f"{PL}$L$5:$O$7"; PHDR=f"{PL}$L$4:$O$4"
PDI=f"{PL}$C$5:$C$7"; PDF=f"{PL}$D$5:$D$7"; POI=f"{PL}$E$5:$E$7"; POF=f"{PL}$F$5:$F$7"
CN="Contributions!"

c3=wb.create_sheet("Cost by Tier")
c3["A1"]="What each Angle plan costs Onyx and the employee"; c3["A1"].font=TTL
c3["A2"]=("Onyx pays the greater of its tier contribution and Angle's minimum funding, never more than the premium. "
          "Worst case = the employee's annual premium plus that household's in-network out-of-pocket maximum; premiums never count "
          "toward an out-of-pocket maximum, so the two add without overlapping.")
c3["A2"].font=SUB
hdr(c3,4,["Plan","Tier","Code","Enrolled","Monthly premium","Onyx contribution $/mo","Angle minimum $/mo",
          "Onyx pays $/mo","Employee pays $/mo","Onyx share","Onyx $/yr (all enrolled)","Employee $/yr (each)",
          "Deductible","OOP max","Worst case $/yr (each)"],h=40)
r=5; CT_FIRST=5
for pi,row in enumerate(ANGLE):
    for tname,tcode,_t in TIERS:
        c3.cell(row=r,column=1,value=row[0]).font=INK
        c3.cell(row=r,column=2,value=tname).font=INK
        c3.cell(row=r,column=3,value=tcode).font=INK; c3.cell(row=r,column=3).alignment=CTR
        f={
         "D":f'=COUNTIFS({CN}$D$22:$D$26,$C{r},{CN}$E$22:$E$26,"Yes")',
         "E":f"=INDEX({PPREM},MATCH($A{r},{PN},0),MATCH($C{r},{PHDR},0))",
         "F":f"=INDEX({CN}$C$7:$C$10,MATCH($C{r},{CN}$B$7:$B$10,0))",
         "G":f"={CN}$C$16",
         "H":f"=MIN($E{r},MAX($F{r},$G{r}))",
         "I":f"=$E{r}-$H{r}",
         "J":f"=IF($E{r}=0,0,$H{r}/$E{r})",
         "K":f"=$H{r}*$D{r}*12",
         "L":f"=$I{r}*12",
         "M":f'=IF($C{r}="EE",INDEX({PDI},MATCH($A{r},{PN},0)),INDEX({PDF},MATCH($A{r},{PN},0)))',
         "N":f'=IF($C{r}="EE",INDEX({POI},MATCH($A{r},{PN},0)),INDEX({POF},MATCH($A{r},{PN},0)))',
         "O":f"=$L{r}+$N{r}",
        }
        for col,formula in f.items():
            cell=c3[f"{col}{r}"]; cell.value=formula
            cell.font=GRN if col in ("E","F","G","M","N") else INK
        for col in ("E","F","G","H","I"): c3[f"{col}{r}"].number_format=CUR2
        for col in ("K","L","M","N","O"): c3[f"{col}{r}"].number_format=CUR
        c3[f"J{r}"].number_format=PCT
        c3[f"D{r}"].alignment=CTR
        c3[f"O{r}"].font=Font(name=F,size=10,bold=True)
        for col in range(1,16): c3.cell(row=r,column=col).border=BOX
        if tcode in ("ES","FAM"):
            for col in range(1,16): c3.cell(row=r,column=col).fill=BAND
        r+=1
CT_LAST=r-1
c3.cell(row=CT_LAST+2,column=1,
  value=("HSA funding is deliberately not modeled. On the HDHP 3400/5000 an employee can pay part of that worst case with pretax "
         "dollars, so the real cost is lower - but by an amount only the employee controls.")).font=SUB
c3.cell(row=CT_LAST+3,column=1,
  value="'Onyx $/yr (all enrolled)' multiplies by how many people sit in that tier, so the column sums to Onyx's whole medical spend.").font=SUB
for col,w in [("A",24),("B",21),("C",7),("D",10),("E",14),("F",16),("G",15),("H",13),("I",15),
              ("J",11),("K",19),("L",17),("M",13),("N",13),("O",18)]:
    c3.column_dimensions[col].width=w
c3.freeze_panes="D5"; c3.sheet_view.showGridLines=False
c3.auto_filter.ref=f"A4:O{CT_LAST}"
print("cost by tier ok", CT_FIRST, CT_LAST)

# ================= 4. BY EMPLOYEE =================
e=wb.create_sheet("By Employee")
e["A1"]="Every Angle plan, priced for each person"; e["A1"].font=TTL
e["A2"]="Driven by the enrollment table on the Contributions tab. Anyone marked No shows a blank deduction - change them to Yes to price them in."
e["A2"].font=SUB
hdr(e,4,["Employee","Enrolling?","Tier","Code","Plan","Monthly premium","Onyx pays $/mo",
         "Employee $/mo","Employee $/yr","Deductible","OOP max","Worst case $/yr"],h=34)
r=5; E_FIRST=5
for ei,(n,pay,tier,code,yn,_note) in enumerate(EMP):
    for row in ANGLE:
        key=f'"{row[0]}"'
        e.cell(row=r,column=1,value=n).font=INK
        e[f"B{r}"]=f"={CN}$E${22+ei}"; e[f"B{r}"].font=GRN; e[f"B{r}"].alignment=CTR
        e.cell(row=r,column=3,value=tier).font=INK
        e.cell(row=r,column=4,value=code).font=INK; e.cell(row=r,column=4).alignment=CTR
        e.cell(row=r,column=5,value=row[0]).font=INK
        m=f'MATCH($E{r}&"|"&$D{r},CostKey,0)'
        for col,src in [("F","E"),("G","H"),("H","I"),("I","L"),("J","M"),("K","N"),("L","O")]:
            e[f"{col}{r}"]=f"=INDEX('Cost by Tier'!${src}${CT_FIRST}:${src}${CT_LAST},{m})"
            e[f"{col}{r}"].font=GRN
        for col in ("F","G","H"): e[f"{col}{r}"].number_format=CUR2
        for col in ("I","J","K","L"): e[f"{col}{r}"].number_format=CUR
        e[f"L{r}"].font=Font(name=F,size=10,bold=True,color="008000")
        for col in range(1,13): e.cell(row=r,column=col).border=BOX
        if yn!="Yes":
            for col in range(1,13): e.cell(row=r,column=col).fill=BAND
        r+=1
E_LAST=r-1
for col,w in [("A",22),("B",11),("C",21),("D",7),("E",24),("F",14),("G",14),("H",14),("I",14),("J",13),("K",13),("L",16)]:
    e.column_dimensions[col].width=w
e.freeze_panes="F5"; e.sheet_view.showGridLines=False
e.auto_filter.ref=f"A4:L{E_LAST}"
# helper key column on Cost by Tier for the lookup above
c3.cell(row=4,column=17,value="Row key").font=HDF
c3.cell(row=4,column=17).fill=HFILL; c3.cell(row=4,column=17).border=BOX
for rr in range(CT_FIRST,CT_LAST+1):
    c3[f"Q{rr}"]=f'=$A{rr}&"|"&$C{rr}'; c3[f"Q{rr}"].font=INK; c3[f"Q{rr}"].border=BOX
c3.column_dimensions["Q"].width=30; c3.column_dimensions["Q"].hidden=True
wb.defined_names.add(openpyxl.workbook.defined_name.DefinedName(
    "CostKey", attr_text=f"'Cost by Tier'!$Q${CT_FIRST}:$Q${CT_LAST}"))
print("by employee ok", E_FIRST, E_LAST)
wb.save("_stage2.xlsx"); print("stage 2 saved")

# --- add an all-enrolled employee-cost column to Cost by Tier (col P) ---
hp=c3.cell(row=4,column=16,value="Employee $/yr (all in tier)")
hp.font=HDF; hp.fill=HFILL; hp.border=BOX
hp.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
for rr in range(CT_FIRST,CT_LAST+1):
    cell=c3[f"P{rr}"]; cell.value=f"=$I{rr}*$D{rr}*12"; cell.font=INK
    cell.number_format=CUR; cell.border=BOX
    if c3[f"C{rr}"].value in ("ES","FAM"): cell.fill=BAND
c3.column_dimensions["P"].width=18
c3.auto_filter.ref=f"A4:P{CT_LAST}"

# ================= 5. QUESTCO =================
q=wb.create_sheet("Questco")
q["A1"]="Questco - PEO services other than medical"; q["A1"].font=TTL
q["A2"]="IRS-Certified PEO. Quote dated for 5 worksite employees on $658,322 of annual payroll. Medical is declined - Onyx takes Angle instead."
q["A2"].font=SUB

sec(q,4,"1.  WHAT ONYX PAYS QUESTCO   (payroll taxes are separate - see block 3)")
hdr(q,5,["Cost line","Basis","Rate / amount","Headcount","Elected?","Monthly","Annual","Note"])
q["D6"]=5; q["D6"].font=BLUB; q["D6"].fill=YEL; q["D6"].alignment=CTR
QR=[("Administrative fee","Per active employee per month",117.00,"Required",
     "The core PEO fee. $117 PEPM as quoted."),
    ("Workers' compensation + EPLI","Annual premium",1402.13,"Required",
     "Subject to full underwriting - loss runs, class-code confirmation, possible survey. The rate can move."),
    ("Cyber liability","Annual premium",420.00,"No",
     "$250,000 aggregate, $1,000 per-claim retention, $50,000 sub-limits. Set to No - Onyx is placing cyber separately. A standalone policy is the right call for a firm holding client tax data under the FTC Safeguards Rule and IRS Pub 4557."),
   ]
for i,(line,basis,rate,el,note) in enumerate(QR):
    r=6+i
    q.cell(row=r,column=1,value=line).font=INK
    q.cell(row=r,column=2,value=basis).font=SM
    c=q.cell(row=r,column=3,value=rate); c.font=BLUB; c.fill=YEL; c.number_format=CUR2
    c=q.cell(row=r,column=5,value=el); c.alignment=CTR
    c.font=SM if el=="Required" else BLUB
    if el!="Required": c.fill=YEL
    if i==0:
        q.cell(row=r,column=7,value="=C6*D6*12").font=INK
    else:
        q.cell(row=r,column=7,value=f'=IF($E{r}="No",0,$C{r})').font=INK
    q.cell(row=r,column=6,value=f"=G{r}/12").font=INK
    q.cell(row=r,column=6).number_format=CUR2
    q.cell(row=r,column=7).number_format=CUR
    q.cell(row=r,column=8,value=note).font=SM; q.cell(row=r,column=8).alignment=WRAP
    q.row_dimensions[r].height=40
    for col in range(1,9): q.cell(row=r,column=col).border=BOX
dvc=DataValidation(type="list",formula1='"Yes,No"',allow_blank=False,showDropDown=False)
q.add_data_validation(dvc); dvc.add("E8")
r=9
q.cell(row=r,column=1,value="Recurring subtotal").font=BLD
q.cell(row=r,column=6,value="=SUM(F6:F8)").font=BLD; q.cell(row=r,column=6).number_format=CUR2
q.cell(row=r,column=7,value="=SUM(G6:G8)").font=BLD; q.cell(row=r,column=7).number_format=CUR
for col in range(1,9):
    q.cell(row=r,column=col).border=BOX; q.cell(row=r,column=col).fill=TOTFILL
q["A10"]="Payroll installation / implementation"; q["A10"].font=INK
q["B10"]="One time, first year only"; q["B10"].font=SM
c=q["G10"]; c.value=1000.00; c.font=BLUB; c.fill=YEL; c.number_format=CUR
for col in range(1,9): q.cell(row=10,column=col).border=BOX
q["A11"]="The 401(k) is on its own tab. Questco charges Onyx nothing for it unless Onyx matches."
q["A11"].font=SUB

sec(q,13,"2.  OPTIONAL EMPLOYER-PAID BENEFITS   <- set Yes / No")
hdr(q,14,["Benefit","Elected?","Design","Monthly","Annual","Note"])
ANC=[("Basic life & AD&D","Yes","$50,000 flat",46.25,
      "$50,000 is exactly the IRC 79 exclusion limit. Going to 1x salary would create imputed W-2 income for Steven, Lisa and Jessica and cost about $1,100 more. Steven's is taxable either way as a >2% shareholder."),
     ("Long-term disability","Yes","60% to $10,000/mo, 90-day elimination",159.09,
      "The $15,000 tier is priced identically because nobody's 60% benefit reaches it - Steven's is the highest at $9,300. At $5,000 the three highest earners would be underinsured."),
     ("Short-term disability","No","60% to $1,500/wk, 7/14, 13 weeks",126.74,
      "Onyx's stated benefit design has STD employee-paid. Set to Yes to model Onyx funding it."),
    ]
for i,(b,el,design,mon,note) in enumerate(ANC):
    r=15+i
    q.cell(row=r,column=1,value=b).font=INK
    c=q.cell(row=r,column=2,value=el); c.font=BLUB; c.fill=YEL; c.alignment=CTR
    q.cell(row=r,column=3,value=design).font=SM
    c=q.cell(row=r,column=4,value=mon); c.font=BLUB; c.fill=YEL; c.number_format=CUR2
    c=q.cell(row=r,column=5,value=f'=IF($B{r}="Yes",$D{r}*12,0)'); c.font=INK; c.number_format=CUR
    q.cell(row=r,column=6,value=note).font=SM; q.cell(row=r,column=6).alignment=WRAP
    q.row_dimensions[r].height=42
    for col in range(1,7): q.cell(row=r,column=col).border=BOX
dv2=DataValidation(type="list",formula1='"Yes,No"',allow_blank=False,showDropDown=False)
q.add_data_validation(dv2); dv2.add("B15:B17")
r=18
q.cell(row=r,column=1,value="Elected ancillary subtotal").font=BLD
q.cell(row=r,column=4,value="=SUM(E15:E17)/12").font=BLD; q.cell(row=r,column=4).number_format=CUR2
q.cell(row=r,column=5,value="=SUM(E15:E17)").font=BLD; q.cell(row=r,column=5).number_format=CUR
for col in range(1,7):
    q.cell(row=r,column=col).border=BOX; q.cell(row=r,column=col).fill=TOTFILL
q["A19"]="Rates are estimates from the Questco ancillary comparison - final rates come from the salaries in payroll. Dental and vision are employee-paid and are not in this model."
q["A19"].font=SUB

sec(q,21,"3.  PAYROLL TAXES   (Onyx owes these with or without a PEO - shown so the invoice reconciles, NOT as a cost of choosing Questco)")
hdr(q,22,["Tax","Quoted annual","Note"])
TAX=[("FICA - employer share",50665.81,"Quoted figure exceeds 7.65% of $658,322 ($50,361.63) and does not appear to cap Social Security at the wage base. With Steven at $186,000 and Lisa at $178,094 it should be lower, not higher. Ask Questco what base was used."),
     ("FUTA",210.00,"5 employees x $7,000 x 0.6%. Ties exactly."),
     ("SUTA (Arizona)",1300.00,"Confirm whether Arizona reporting is at client level or PEO level - it decides whether Onyx keeps its own experience rating on exit.")]
for i,(t,amt,note) in enumerate(TAX):
    r=23+i
    q.cell(row=r,column=1,value=t).font=INK
    c=q.cell(row=r,column=2,value=amt); c.font=BLU; c.number_format=CUR
    q.cell(row=r,column=3,value=note).font=SM; q.cell(row=r,column=3).alignment=WRAP
    q.row_dimensions[r].height=40
    for col in range(1,4): q.cell(row=r,column=col).border=BOX
r=26
q.cell(row=r,column=1,value="Pass-through subtotal").font=BLD
q.cell(row=r,column=2,value="=SUM(B23:B25)").font=BLD; q.cell(row=r,column=2).number_format=CUR
for col in range(1,4):
    q.cell(row=r,column=col).border=BOX; q.cell(row=r,column=col).fill=TOTFILL

sec(q,28,"4.  WHAT THE FEE BUYS")
hdr(q,29,["Category","Included","Relevance to Onyx"])
SVC=[("Payroll administration","Processing, direct deposit, garnishments, state and federal tax remittance, Forms 940/941, W-2/W-3, PTO tracking, ACA reporting, 20+ standard reports. Payroll runs 3 business days before payday and Onyx controls entry.","Core. Replaces the payroll function outright."),
     ("HR expertise and support","Dedicated HR team, handbook and policy review, job descriptions, onboarding, employee relations, terminations, unemployment claims, EEOC, FMLA and FLSA compliance.","High. Displaces the outside HR support Onyx pays for today."),
     ("Benefits administration","Enrollments, terminations, qualifying events, COBRA, Section 125 compliance, creditable coverage reporting, claims advocacy, and in-house invoice reconciliation for client-owned plans.","High - and the reason Questco can administer the Angle plan."),
     ("Self-funded plan filings","PCORI fee calculation and Form 720; Forms 1094-B/1095-B.","CRITICAL. Angle is level funded, so these are now Onyx's obligations. Confirm in writing that Questco covers them for the Angle plan by name."),
     ("Workers' compensation","Coverage at PEO rates, safety training and materials, claims handling.","Medium. Office exposure is low, but the rate is still subject to underwriting."),
     ("401(k)","Large-group multiple employer plan, $3.00 per participating employee per quarter, no employer cost unless Onyx matches. Confirmed it does not require annual re-signature by Onyx.","High. Cheapest participant fee of any vendor reviewed. See the 401(k) tab."),
     ("Cyber liability","$250,000 annual aggregate, $1,000 per-claim retention, $50,000 sub-limits, 24/7 incident hotline.","Declined - Onyx is placing a standalone policy instead."),
     ("CPEO status","IRS-Certified Professional Employer Organization.","High. It is what protects the payroll-tax wage bases from restarting on a mid-year start - worth roughly $5,600 given Steven's and Lisa's salaries. Confirm it."),
     ("HR technology","Employee self-service portal, mobile access, reporting and analytics.","Medium."),
    ]
for i,(cat,inc,rel) in enumerate(SVC):
    r=30+i
    q.cell(row=r,column=1,value=cat).font=BLD
    q.cell(row=r,column=2,value=inc).font=SM; q.cell(row=r,column=2).alignment=WRAP
    q.cell(row=r,column=3,value=rel).font=SM; q.cell(row=r,column=3).alignment=WRAP
    q.row_dimensions[r].height=46
    for col in range(1,4): q.cell(row=r,column=col).border=BOX
for col,w in [("A",30),("B",42),("C",30),("D",12),("E",12),("F",14),("G",16),("H",56)]:
    q.column_dimensions[col].width=w
q.sheet_view.showGridLines=False
print("questco ok")

# ================= 6. 401(k) =================
k=wb.create_sheet("401(k)")
k["A1"]="401(k) through Questco - match modelling"; k["A1"].font=TTL
k["A2"]=("Questco charges $3.00 per participating employee per quarter and nothing to Onyx unless Onyx matches. "
         "Everything below is a lever - change the match formula, who participates, or what they defer.")
k["A2"].font=SUB
k["A3"]="EDIT THE YELLOW CELLS."; k["A3"].font=Font(name=F,size=10,bold=True,color="8E2F2A")

sec(k,5,"1.  PLAN SETTINGS",5)
hdr(k,6,["Setting","","Value","Per / of pay","Note"])
def setrow(r,label,v1,v2,note,f1=CUR2,f2=None,yellow2=False):
    k.cell(row=r,column=1,value=label).font=INK
    c=k.cell(row=r,column=3,value=v1); c.font=BLUB; c.fill=YEL
    if f1: c.number_format=f1
    if v2 is not None:
        c=k.cell(row=r,column=4,value=v2)
        if yellow2: c.font=BLUB; c.fill=YEL
        else: c.font=SM
        if f2: c.number_format=f2
    k.cell(row=r,column=5,value=note).font=SM; k.cell(row=r,column=5).alignment=WRAP
    k.row_dimensions[r].height=30
    for col in range(1,6): k.cell(row=r,column=col).border=BOX

setrow(7,"Per-participant fee",3.00,"per quarter",
       "Questco's quoted 401(k) fee. $12 a year per participant - the lowest of any vendor reviewed. TriNet charges $63; ADP charges none but recovers it inside fund expense ratios.")
setrow(8,"Fee paid by","Participant",None,
       "Set to Employer to have Onyx absorb it. Questco's statement is that Onyx pays nothing unless it matches, so Participant is the default - confirm with them.",f1=None)
setrow(9,"IRS compensation limit",350000,None,
       "CONFIRM the current-year 401(a)(17) figure. Nobody here is close, so it does not bind.",f1=CUR)
setrow(10,"IRS elective deferral limit",24000,None,
       "CONFIRM the current-year 402(g) figure. Not binding at the deferral rates below.",f1=CUR)
dvf=DataValidation(type="list",formula1='"Participant,Employer"',allow_blank=False,showDropDown=False)
k.add_data_validation(dvf); dvf.add("C8")

setrow(12,"Match type","Tiered match",None,
       "Tiered match, Non-elective, or None. Non-elective goes to every participating employee whether or not they defer.",f1=None)
setrow(13,"Tier 1 - match rate",1.00,0.03,
       "Match 100% of the first 3% of pay deferred. With Tier 2 this is the classic safe harbor basic formula.",f1=PCT,f2=PCT,yellow2=True)
setrow(14,"Tier 2 - match rate",0.50,0.02,
       "Match 50% of the next 2% of pay. An employee deferring 5% or more earns the full 4% of pay.",f1=PCT,f2=PCT,yellow2=True)
setrow(15,"Non-elective rate",0.03,None,
       "Used only when Match type is Non-elective. 3% of pay to every participating employee.",f1=PCT)
dvm=DataValidation(type="list",formula1='"Tiered match,Non-elective,None"',allow_blank=False,showDropDown=False)
k.add_data_validation(dvm); dvm.add("C12")

sec(k,17,"2.  WHO PARTICIPATES AND WHAT THEY DEFER",8)
hdr(k,18,["Employee","Eligible pay","Participating?","Deferral % of pay","Employee deferral $/yr",
          "Onyx match $/yr","Match as % of pay","Participant fee $/yr"],h=34)
K_FIRST=19
for i,(n,pay,tier,code,yn,_note) in enumerate(EMP):
    r=K_FIRST+i
    k.cell(row=r,column=1,value=n).font=INK
    c=k.cell(row=r,column=2,value=pay); c.font=BLU; c.number_format=CUR
    c=k.cell(row=r,column=3,value="Yes"); c.font=BLUB; c.fill=YEL; c.alignment=CTR
    c=k.cell(row=r,column=4,value=0.06); c.font=BLUB; c.fill=YEL; c.number_format=PCT
    c=k.cell(row=r,column=5,value=f'=IF($C{r}<>"Yes",0,MIN(MIN($B{r},$C$9)*$D{r},$C$10))')
    c.font=INK; c.number_format=CUR
    c=k.cell(row=r,column=6,value=(
        f'=IF($C{r}<>"Yes",0,'
        f'IF($C$12="None",0,'
        f'IF($C$12="Non-elective",MIN($B{r},$C$9)*$C$15,'
        f'MIN($B{r},$C$9)*($C$13*MIN($D{r},$D$13)+$C$14*MIN(MAX($D{r}-$D$13,0),$D$14)))))'))
    c.font=INK; c.number_format=CUR
    c=k.cell(row=r,column=7,value=f"=IF($B{r}=0,0,$F{r}/$B{r})"); c.font=INK; c.number_format=PCT
    c=k.cell(row=r,column=8,value=f'=IF($C{r}<>"Yes",0,$C$7*4)'); c.font=INK; c.number_format=CUR
    for col in range(1,9): k.cell(row=r,column=col).border=BOX
K_LAST=K_FIRST+len(EMP)-1
dvp=DataValidation(type="list",formula1='"Yes,No"',allow_blank=False,showDropDown=False)
k.add_data_validation(dvp); dvp.add(f"C{K_FIRST}:C{K_LAST}")
r=K_LAST+1
k.cell(row=r,column=1,value="TOTAL").font=BLD
k.cell(row=r,column=2,value=f"=SUM(B{K_FIRST}:B{K_LAST})").font=BLD
k.cell(row=r,column=2).number_format=CUR
k.cell(row=r,column=3,value=f'=COUNTIF(C{K_FIRST}:C{K_LAST},"Yes")&" in"').font=BLD
k.cell(row=r,column=3).alignment=CTR
for col,letter in ((5,"E"),(6,"F"),(8,"H")):
    c=k.cell(row=r,column=col,value=f"=SUM({letter}{K_FIRST}:{letter}{K_LAST})")
    c.font=BLD; c.number_format=CUR
c=k.cell(row=r,column=7,value=f"=IF(B{r}=0,0,F{r}/B{r})"); c.font=BLD; c.number_format=PCT
for col in range(1,9):
    k.cell(row=r,column=col).border=BOX; k.cell(row=r,column=col).fill=TOTFILL
K_TOT=r
k.cell(row=K_TOT+1,column=1,
  value="Deferral rates default to 6%, which fully earns the safe harbor basic match. They are an assumption - nobody has elected yet. ADP's own multiple employer plan reports a 7.81% average.").font=SUB

sec(k,K_TOT+3,"3.  WHAT THE 401(k) COSTS ONYX",5)
hdr(k,K_TOT+4,["Line","Monthly","Annual","Note"])
KS=[("Employer match",f"=C{K_TOT+5}/12",f"=F{K_TOT}",
     "Deductible to Onyx. Unlike health premiums, a match for Steven as a >2% S-corp shareholder is a normal deductible plan contribution, not W-2 wages."),
    ("Participant fees absorbed by Onyx",f"=C{K_TOT+6}/12",f'=IF($C$8="Employer",H{K_TOT},0)',
     "Zero while the fee sits on participants. Flip 'Fee paid by' to Employer to model Onyx absorbing it."),
   ]
for i,(line,mon,ann,note) in enumerate(KS):
    r=K_TOT+5+i
    k.cell(row=r,column=1,value=line).font=INK
    c=k.cell(row=r,column=2,value=f"=C{r}/12"); c.font=INK; c.number_format=CUR2
    c=k.cell(row=r,column=3,value=ann); c.font=INK; c.number_format=CUR
    k.cell(row=r,column=4,value=note).font=SM; k.cell(row=r,column=4).alignment=WRAP
    k.row_dimensions[r].height=34
    for col in range(1,5): k.cell(row=r,column=col).border=BOX
r=K_TOT+7
k.cell(row=r,column=1,value="ONYX 401(k) COST").font=BLD
c=k.cell(row=r,column=2,value=f"=SUM(B{K_TOT+5}:B{K_TOT+6})"); c.font=BLD; c.number_format=CUR2
c=k.cell(row=r,column=3,value=f"=SUM(C{K_TOT+5}:C{K_TOT+6})"); c.font=BLD; c.number_format=CUR
for col in range(1,5):
    k.cell(row=r,column=col).border=BOX; k.cell(row=r,column=col).fill=TOTFILL
K_COST=r
k.cell(row=r+1,column=1,value=f"Participants pay $0 while the fee stays on them; the total they bear is on row {K_TOT} column H.").font=SUB

sec(k,K_COST+3,"4.  BEFORE ADOPTING",5)
for i,t in enumerate([
  "Steven owns 100%, so the plan is almost certainly top-heavy. A safe harbor design - the tiered match above, or the 3% non-elective - exempts it from the top-heavy minimum. A discretionary match does not.",
  "Ask Questco whether an adopting employer in their multiple employer plan still claims the SECURE 2.0 start-up credits: up to $5,000/yr for three years, $500/yr for auto-enrolment. Neither Questco nor any other vendor reviewed has answered it.",
  "Ask for the fund lineup with expense ratios and the target-date series. The $3 per quarter is only the visible fee; the expense ratios are where the rest sits.",
  "Confirm whether the $3 is billed to Onyx or deducted from participant accounts, and whether the match triggers Questco's $150/hour allocation work or anything like it.",
  "Confirm the 401(k) does not require annual re-signature by Onyx - the answer that moved this decision forward in the first place.",
]):
    k.cell(row=K_COST+4+i,column=1,value=u"•  "+t).font=INK
    k.cell(row=K_COST+4+i,column=1).alignment=WRAP
    k.row_dimensions[K_COST+4+i].height=28
for col,w in [("A",40),("B",16),("C",18),("D",60),("E",62),("F",16),("G",17),("H",18)]:
    k.column_dimensions[col].width=w
k.sheet_view.showGridLines=False
print("401k ok  rows", K_FIRST, K_LAST, "cost row", K_COST)

# ================= 7. COMBINED =================
CB=wb.create_sheet("Combined")
CT="'Cost by Tier'!"
PLAN_A=f"{CT}$A${CT_FIRST}:$A${CT_LAST}"
ONYX_K=f"{CT}$K${CT_FIRST}:$K${CT_LAST}"
EMP_P=f"{CT}$P${CT_FIRST}:$P${CT_LAST}"
WORST_O=f"{CT}$O${CT_FIRST}:$O${CT_LAST}"
ENR_D=f"{CT}$D${CT_FIRST}:$D${CT_LAST}"
QO="Questco!"; KO="'401(k)'!"
NONMED_M=f"{QO}$F$6+{QO}$F$7+{QO}$F$8+{QO}$D$18+{KO}$B${K_COST}"
NONMED_A=f"{QO}$G$6+{QO}$G$7+{QO}$G$8+{QO}$E$18+{KO}$C${K_COST}"

CB["A1"]="Angle + Questco - monthly and annual projection"; CB["A1"].font=TTL
CB["A2"]="Everything here recalculates from the Contributions, Questco and 401(k) tabs. Pick a plan below to drive the first two blocks."
CB["A2"].font=SUB
CB["A4"]="SELECTED ANGLE PLAN"; CB["A4"].font=BLD
c=CB["C4"]; c.value=ANGLE[0][0]; c.font=BLUB; c.fill=YEL
dv3=DataValidation(type="list",formula1='"{}"'.format(",".join(r[0] for r in ANGLE)),
                   allow_blank=False,showDropDown=False)
CB.add_data_validation(dv3); dv3.add("C4")
for col in range(1,4): CB.cell(row=4,column=col).border=BOX

sec(CB,6,"1.  WHAT ONYX PAYS",6)
hdr(CB,7,["Cost line","Vendor","Monthly","Annual","Note"])
LINES=[("Medical - employer contribution","Angle / CapFi",
        f"=SUMIFS({ONYX_K},{PLAN_A},$C$4)/12",f"=SUMIFS({ONYX_K},{PLAN_A},$C$4)",
        "Fixed dollars per tier. Does not move when an employee changes plans."),
       ("Administrative fee","Questco",f"={QO}F6",f"={QO}G6","$117 per employee per month."),
       ("Workers' comp + EPLI","Questco",f"={QO}F7",f"={QO}G7","Subject to underwriting."),
       ("Cyber liability","Questco",f"={QO}F8",f"={QO}G8",
        "Zero while it is set to No on the Questco tab. Onyx is placing this separately."),
       ("Employer-paid ancillary","Questco",f"={QO}D18",f"={QO}E18",
        "Life, LTD and STD as elected on the Questco tab."),
       ("401(k) - employer cost","Questco",f"={KO}B{K_COST}",f"={KO}C{K_COST}",
        "Match, plus participant fees only if Onyx absorbs them. Set the formula on the 401(k) tab."),
      ]
for i,(line,ven,mon,ann,note) in enumerate(LINES):
    r=8+i
    CB.cell(row=r,column=1,value=line).font=INK
    CB.cell(row=r,column=2,value=ven).font=SM
    c=CB.cell(row=r,column=3,value=mon); c.font=INK; c.number_format=CUR2
    c=CB.cell(row=r,column=4,value=ann); c.font=INK; c.number_format=CUR
    CB.cell(row=r,column=5,value=note).font=SM; CB.cell(row=r,column=5).alignment=WRAP
    CB.row_dimensions[r].height=28
    for col in range(1,6): CB.cell(row=r,column=col).border=BOX
CB["A14"]="ONYX RECURRING TOTAL"; CB["A14"].font=BLD
CB["C14"]="=SUM(C8:C13)"; CB["C14"].font=BLD; CB["C14"].number_format=CUR2
CB["D14"]="=SUM(D8:D13)"; CB["D14"].font=BLD; CB["D14"].number_format=CUR
CB["A15"]="Implementation (first year only)"; CB["A15"].font=INK
CB["B15"]="Questco"; CB["B15"].font=SM
CB["D15"]=f"={QO}G10"; CB["D15"].font=INK; CB["D15"].number_format=CUR
CB["A16"]="ONYX FIRST-YEAR TOTAL"; CB["A16"].font=BLD
CB["D16"]="=D14+D15"; CB["D16"].font=BLD; CB["D16"].number_format=CUR
for r2 in (14,15,16):
    for col in range(1,6): CB.cell(row=r2,column=col).border=BOX
for r2 in (14,16):
    for col in range(1,6): CB.cell(row=r2,column=col).fill=TOTFILL
CB["A17"]=("The 401(k) line is a MODELLED match, not a committed cost. At the default settings - safe harbor basic, everyone "
           "deferring 6% - it is the largest line on this page after medical. Set Match type to None on the 401(k) tab to see Onyx's "
           "cost without a match.")
CB["A17"].font=Font(name=F,size=9,bold=True,italic=True,color="8E2F2A")
CB["A18"]=("Payroll taxes of about $52,176 a year are NOT in this total. Onyx owes FICA, FUTA and SUTA whether or not there is a PEO, "
           "so they are a pass-through on the Questco invoice rather than a cost of the relationship. They are itemised on the Questco tab.")
CB["A18"].font=SUB

sec(CB,20,"2.  WHAT THE EMPLOYEES PAY, AND THE COMBINED PICTURE",6)
hdr(CB,21,["Line","Monthly","Annual","Note"])
EL=[("Employee medical deductions, all enrolled",f"=SUMIFS({EMP_P},{PLAN_A},$C$4)/12",
     f"=SUMIFS({EMP_P},{PLAN_A},$C$4)","Premium only. Pretax for everyone except Steven."),
    ("Employee 401(k) deferrals",f"={KO}E{K_TOT}/12",f"={KO}E{K_TOT}",
     "Their own savings, not a cost - shown so the payroll picture is complete."),
    ("Employee 401(k) fees",f"=IF({KO}$C$8=\"Employer\",0,{KO}H{K_TOT}/12)",
     f"=IF({KO}$C$8=\"Employer\",0,{KO}H{K_TOT})","$3 per quarter each while the fee sits on participants."),
    ("Onyx + employees, medical premium only","=C8+B22","=D8+C22",
     "The true cost of the Angle plan itself, both sides."),
    ("Onyx + employees, all in","=C14+B22+B24","=D14+C22+C24",
     "Adds the Questco fee, elected ancillary, the 401(k) and the participant fees."),
    ("Combined worst case, one bad year","",
     f"=D14+SUMPRODUCT(({PLAN_A}=$C$4)*{WORST_O}*{ENR_D})",
     "Assumes every enrolled household hits its out-of-pocket maximum in the same year. A ceiling, not a forecast."),
   ]
for i,(line,mon,ann,note) in enumerate(EL):
    r=22+i
    CB.cell(row=r,column=1,value=line).font=BLD if i>=3 else INK
    if mon:
        c=CB.cell(row=r,column=2,value=mon); c.font=INK; c.number_format=CUR2
    c=CB.cell(row=r,column=3,value=ann); c.font=BLD if i>=3 else INK; c.number_format=CUR
    CB.cell(row=r,column=4,value=note).font=SM; CB.cell(row=r,column=4).alignment=WRAP
    CB.row_dimensions[r].height=28
    for col in range(1,5): CB.cell(row=r,column=col).border=BOX

sec(CB,29,"3.  ALL THREE PLANS SIDE BY SIDE   (annual, at the current settings)",7)
hdr(CB,30,["Angle plan","HSA","Onyx - medical","Onyx - all in","Employee medical","Onyx + employees","Combined worst case"])
for i,row in enumerate(ANGLE):
    r=31+i
    CB.cell(row=r,column=1,value=row[0]).font=INK
    CB.cell(row=r,column=2,value=row[1]).font=INK; CB.cell(row=r,column=2).alignment=CTR
    CB[f"C{r}"]=f"=SUMIFS({ONYX_K},{PLAN_A},$A{r})"
    CB[f"D{r}"]=f"=C{r}+{NONMED_A}"
    CB[f"E{r}"]=f"=SUMIFS({EMP_P},{PLAN_A},$A{r})"
    CB[f"F{r}"]=f'=D{r}+E{r}+IF({KO}$C$8="Employer",0,{KO}H{K_TOT})'
    CB[f"G{r}"]=f"=D{r}+SUMPRODUCT(({PLAN_A}=$A{r})*{WORST_O}*{ENR_D})"
    for col in "CDEFG":
        CB[f"{col}{r}"].font=INK; CB[f"{col}{r}"].number_format=CUR
    CB[f"D{r}"].font=BLD
    for col in range(1,8): CB.cell(row=r,column=col).border=BOX
    if i%2:
        for col in range(1,8): CB.cell(row=r,column=col).fill=BAND
CB["A35"]=("Onyx's medical column is the same across all three plans because the contribution is a fixed dollar amount and Angle's minimum "
           "funding never binds at these settings. A richer plan costs Onyx nothing extra and costs the employee the whole difference. "
           "Lower the contribution far enough on the Contributions tab and the minimum starts to bind - the check column there says so.")
CB["A35"].font=SUB
CB["A36"]="Angle offers no dental or vision. Questco quotes both as employee-paid, so neither appears in Onyx's cost here."
CB["A36"].font=SUB
for col,w in [("A",42),("B",18),("C",18),("D",18),("E",18),("F",20),("G",22)]:
    CB.column_dimensions[col].width=w
CB.sheet_view.showGridLines=False
print("combined ok")

import os
for tmp in ("_stage1.xlsx","_stage2.xlsx","_stage3.xlsx","_head.py"):
    if os.path.exists(tmp): os.remove(tmp)
wb.save("Onyx_Angle_Questco_Cost_Model.xlsx")
print("SAVED")
