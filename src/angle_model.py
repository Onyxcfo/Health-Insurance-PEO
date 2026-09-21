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
GREY=Font(name=F,size=10,color="808080"); GREYI=Font(name=F,size=9,italic=True,color="808080")
HDF=Font(name=F,size=9,bold=True,color="FFFFFF"); SECF=Font(name=F,size=11,bold=True,color="1D5B72")
HFILL=PatternFill("solid",fgColor="1D5B72"); YEL=PatternFill("solid",fgColor="FFFF00")
BAND=PatternFill("solid",fgColor="EFF3F5"); TOTFILL=PatternFill("solid",fgColor="E4E9EA")
OFFFILL=PatternFill("solid",fgColor="F2F2F2")
thin=Side(style="thin",color="BFBFBF"); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
BOT=Border(bottom=Side(style="thin",color="1D5B72"))
CUR='$#,##0;($#,##0);-'; CUR2='$#,##0.00;($#,##0.00);-'; PCT='0.0%'
WRAP=Alignment(wrap_text=True,vertical="top"); CTR=Alignment(horizontal="center")

# plan, hsa, ded_i, ded_f, oop_i, oop_f, plan_pays, pcp, spec, rx, er, EE, ES, EC, FAM, note
ANGLE=[
 ("ANG HDHP 3400/5000","Y",3400,6800,5000,10000,0.80,"20% after ded","20% after ded","20% after ded (all tiers)","20% after ded",
  439.03,921.97,834.16,1361.00,"CapFi's recommended base plan. Sets the minimum-funding floor."),
 ("ANG TRAD 2000/4000","N",2000,4000,4000,8000,0.80,"$20","$50","$20 / $60","$250",
  492.11,1033.43,934.01,1525.54,"Tier rates DERIVED from the ratios in the two documented plans. Confirm with CapFi before quoting to employees."),
 ("ANG TRAD 1000/2000","N",1000,2000,2000,4000,0.80,"$10","$30","$10 / $30","$200",
  546.33,1147.29,1038.02,1693.61,"Best plan design in the whole field - $1,000 deductible, $2,000 out-of-pocket maximum."),
]
# tier, code, fixed-dollar alternative, include in totals?
TIERS=[("Employee only","EE",300.0,"Yes"),("Employee + spouse","ES",600.0,"Yes"),
       ("Employee + child(ren)","EC",600.0,"No"),("Family","FAM",1000.0,"Yes")]
PCOL={"EE":"L","ES":"M","EC":"N","FAM":"O"}
PCT_DEFAULT=0.50
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

sec(a,5,"1.  HOW ONYX FUNDS THE PREMIUM   <- the master switch")
a["A6"]="Contribution basis"; a["A6"].font=BLD
c=a["C6"]; c.value="Percent of premium"; c.font=BLUB; c.fill=YEL
dvb=DataValidation(type="list",formula1='"Percent of premium,Fixed dollars"',allow_blank=False,showDropDown=False)
a.add_data_validation(dvb); dvb.add("C6")
for col in range(1,4): a.cell(row=6,column=col).border=BOX
a["A7"]=("On 'Percent of premium' Onyx pays the percentage in column D of whatever plan the employee elects, so Onyx's cost moves with "
         "the plan. On 'Fixed dollars' Onyx pays the flat amount in column E and the employee absorbs the whole difference between plans. "
         "Either way Onyx never pays less than Angle's minimum funding, and never more than the premium itself.")
a["A7"].font=SUB

sec(a,9,"2.  ONYX CONTRIBUTION BY TIER   <- the lever")
hdr(a,10,["Tier","Code","Include in totals?","Onyx % of premium","Onyx fixed $/mo (used only on the Fixed dollars basis)",
          "Angle minimum $/mo","Lowest premium in this tier","Status at the current basis"],h=44)
for i,(name,code,amt,inc) in enumerate(TIERS):
    r=11+i; pc=PCOL[code]
    a.cell(row=r,column=1,value=name).font=INK if inc=="Yes" else GREY
    a.cell(row=r,column=2,value=code).font=INK if inc=="Yes" else GREY
    a.cell(row=r,column=2).alignment=CTR
    c=a.cell(row=r,column=3,value=inc); c.font=BLUB; c.fill=YEL; c.alignment=CTR
    c=a.cell(row=r,column=4,value=PCT_DEFAULT); c.font=BLUB; c.fill=YEL; c.number_format=PCT
    c=a.cell(row=r,column=5,value=amt); c.font=BLUB; c.fill=YEL; c.number_format=CUR2
    c=a.cell(row=r,column=6,value="=$C$20"); c.font=GRN; c.number_format=CUR2
    c=a.cell(row=r,column=7,value=f"=MIN('Angle Plans'!${pc}$5:${pc}$7)"); c.font=GRN; c.number_format=CUR2
    c=a.cell(row=r,column=8,value=(
        f'=IF($C{r}<>"Yes","Excluded - carried for history, not counted anywhere",'
        f'IF($C$6="Fixed dollars",IF($E{r}>=$F{r},"Funded at the fixed amount","Topped up to the Angle minimum"),'
        f'IF($D{r}*$G{r}>=$F{r},"Funded at the percentage","Topped up to the Angle minimum on the cheapest plan")))'))
    c.font=SM; c.alignment=WRAP
    a.row_dimensions[r].height=26
    for col in range(1,9): a.cell(row=r,column=col).border=BOX
    if inc!="Yes":
        for col in range(1,9): a.cell(row=r,column=col).fill=OFFFILL
        a.cell(row=r,column=3).fill=YEL
dvi=DataValidation(type="list",formula1='"Yes,No"',allow_blank=False,showDropDown=False)
a.add_data_validation(dvi); dvi.add("C11:C14")
a["A15"]=("Employee + child(ren) is set to No because nobody is expected to elect it. Its rates and economics stay on every tab for the "
          "record, but it is excluded from every total for Onyx and for employees. Switch column C to Yes if that changes.")
a["A15"].font=Font(name=F,size=9,italic=True,color="8E2F2A")
a["A16"]="Medical only. Angle offers no dental or vision; life, LTD and STD are quoted separately by Questco on the Questco tab."
a["A16"].font=SUB

sec(a,18,"3.  ANGLE MINIMUM FUNDING REQUIREMENT",3)
a["A19"]="Percentage of the lowest employee-only rate"; a["A19"].font=INK
c=a["C19"]; c.value=0.50; c.font=BLUB; c.fill=YEL; c.number_format=PCT
a["A20"]="Minimum Onyx must fund, per enrolled employee, per month"; a["A20"].font=BLD
c=a["C20"]; c.value="=C19*MIN('Angle Plans'!L5:L7)"; c.font=BLD; c.number_format=CUR2
a["A21"]="Minimum total funding at current enrollment, per year"; a["A21"].font=INK
c=a["C21"]; c.value="=C20*D40*12"; c.font=INK; c.number_format=CUR
a["A22"]=("Angle's own presentation states minimum funding of $10,536.96/yr - 50% of the HDHP 3400/5000's $439.03 employee-only "
          "rate across four enrolled employees. The formula above lands at $10,536.72; the 24-cent difference is Angle rounding the "
          "monthly floor to $219.52.")
a["A22"].font=SUB
for r in (19,20,21):
    for col in range(1,4): a.cell(row=r,column=col).border=BOX
MINCELL="Contributions!$C$20"
PCTRNG="Contributions!$D$11:$D$14"; FIXRNG="Contributions!$E$11:$E$14"
INCRNG="Contributions!$C$11:$C$14"; CODERNG="Contributions!$B$11:$B$14"
BASIS="Contributions!$C$6"

sec(a,24,"4.  WHO IS ENROLLED   <- change Yes/No to re-run the whole model")
hdr(a,25,["Employee","Annual pay","Medical tier","Code","Enrolling?","","","Note"])
EMP_FIRST=26; EMP_LAST=26+len(EMP)-1
for i,(n,pay,tier,code,yn,note) in enumerate(EMP):
    r=EMP_FIRST+i
    a.cell(row=r,column=1,value=n).font=INK
    c=a.cell(row=r,column=2,value=pay); c.font=BLU; c.number_format=CUR
    a.cell(row=r,column=3,value=tier).font=INK
    a.cell(row=r,column=4,value=code).font=BLU; a.cell(row=r,column=4).alignment=CTR
    c=a.cell(row=r,column=5,value=yn); c.font=BLUB; c.fill=YEL; c.alignment=CTR
    a.cell(row=r,column=8,value=note).font=SM; a.cell(row=r,column=8).alignment=WRAP
    a.row_dimensions[r].height=26
    for col in list(range(1,6))+[8]: a.cell(row=r,column=col).border=BOX
dv=DataValidation(type="list",formula1='"Yes,No"',allow_blank=False,showDropDown=False)
a.add_data_validation(dv); dv.add(f"E{EMP_FIRST}:E{EMP_LAST}")
ENR_CODE=f"Contributions!$D${EMP_FIRST}:$D${EMP_LAST}"
ENR_YN=f"Contributions!$E${EMP_FIRST}:$E${EMP_LAST}"

sec(a,33,"5.  ENROLLMENT AT THE CURRENT SETTINGS",5)
hdr(a,34,["Tier","Code","Include in totals?","Enrolled","Counted in totals"])
for i,(name,code,_amt,_inc) in enumerate(TIERS):
    r=35+i
    a.cell(row=r,column=1,value=name).font=INK
    a.cell(row=r,column=2,value=code).font=INK; a.cell(row=r,column=2).alignment=CTR
    c=a.cell(row=r,column=3,value=f"=$C${11+i}"); c.font=GRN; c.alignment=CTR
    c=a.cell(row=r,column=4,value=f'=COUNTIFS({ENR_CODE},$B{r},{ENR_YN},"Yes")'); c.font=INK; c.alignment=CTR
    c=a.cell(row=r,column=5,value=f'=IF($C{r}="Yes",$D{r},0)'); c.font=INK; c.alignment=CTR
    for col in range(1,6): a.cell(row=r,column=col).border=BOX
    if TIERS[i][3]!="Yes":
        for col in range(1,6): a.cell(row=r,column=col).fill=OFFFILL
r=39
a.cell(row=r,column=1,value="TOTAL").font=BLD
c=a.cell(row=r,column=4,value="=SUM(D35:D38)"); c.font=BLD; c.alignment=CTR
c=a.cell(row=r,column=5,value="=SUM(E35:E38)"); c.font=BLD; c.alignment=CTR
for col in range(1,6):
    a.cell(row=r,column=col).border=BOX; a.cell(row=r,column=col).fill=TOTFILL
a["D40"]="=D39"; a["D40"].font=Font(name=F,size=9,color="FFFFFF")
a["A41"]=("Dollar totals now live on the Elections tab, because on the percentage basis Onyx's cost depends on which plan each person "
          "actually elects. This block counts heads only.")
a["A41"].font=SUB
for col,w in [("A",44),("B",10),("C",17),("D",16),("E",20),("F",16),("G",20),("H",54)]:
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
p["A10"]="The EC column is retained for the record. It is excluded from every total - see column C of the Contributions tab."
p["A10"].font=GREYI
for col,w in [("A",24),("B",6),("C",13),("D",13),("E",15),("F",15),("G",12),("H",15),("I",15),
              ("J",22),("K",15),("L",11),("M",11),("N",11),("O",11),("P",50)]:
    p.column_dimensions[col].width=w
p.freeze_panes="B5"; p.sheet_view.showGridLines=False
print("plans ok")

# ================= 3. COST BY TIER =================
PL="'Angle Plans'!"
PN=f"{PL}$A$5:$A$7"; PPREM=f"{PL}$L$5:$O$7"; PHDR=f"{PL}$L$4:$O$4"
PDI=f"{PL}$C$5:$C$7"; PDF=f"{PL}$D$5:$D$7"; POI=f"{PL}$E$5:$E$7"; POF=f"{PL}$F$5:$F$7"

c3=wb.create_sheet("Cost by Tier")
c3["A1"]="What each Angle plan costs Onyx and the employee"; c3["A1"].font=TTL
c3["A2"]=("Onyx pays the greater of its tier contribution and Angle's minimum funding, never more than the premium. "
          "Worst case = the employee's annual premium plus that household's in-network out-of-pocket maximum; premiums never count "
          "toward an out-of-pocket maximum, so the two add without overlapping.")
c3["A2"].font=SUB
c3["A3"]=("Rows whose tier is excluded on the Contributions tab keep their per-person economics for the record but carry zero enrolled, "
          "so they add nothing to any total.")
c3["A3"].font=GREYI
hdr(c3,4,["Plan","Tier","Code","In totals?","Enrolled","Monthly premium","Onyx contribution $/mo","Angle minimum $/mo",
          "Onyx pays $/mo (each)","Employee pays $/mo (each)","Onyx share","Onyx $/mo (all enrolled)","Onyx $/yr (all enrolled)",
          "Employee $/mo (all in tier)","Employee $/yr (all in tier)","Employee $/yr (each)",
          "Deductible","OOP max","Worst case $/yr (each)"],h=48)
r=5; CT_FIRST=5; CT_TOTROWS=[]
for pi,row in enumerate(ANGLE):
    grp_first=r
    for tname,tcode,_t,_inc in TIERS:
        c3.cell(row=r,column=1,value=row[0]).font=INK
        c3.cell(row=r,column=2,value=tname).font=INK
        c3.cell(row=r,column=3,value=tcode).font=INK; c3.cell(row=r,column=3).alignment=CTR
        f={
         "D":f"=INDEX({INCRNG},MATCH($C{r},{CODERNG},0))",
         "E":f'=IF($D{r}<>"Yes",0,COUNTIFS({ENR_CODE},$C{r},{ENR_YN},"Yes"))',
         "F":f"=INDEX({PPREM},MATCH($A{r},{PN},0),MATCH($C{r},{PHDR},0))",
         "G":(f'=IF({BASIS}="Fixed dollars",INDEX({FIXRNG},MATCH($C{r},{CODERNG},0)),'
              f"INDEX({PCTRNG},MATCH($C{r},{CODERNG},0))*$F{r})"),
         "H":f"={MINCELL}",
         "I":f"=MIN($F{r},MAX($G{r},$H{r}))",
         "J":f"=$F{r}-$I{r}",
         "K":f"=IF($F{r}=0,0,$I{r}/$F{r})",
         "L":f"=$I{r}*$E{r}",
         "M":f"=$L{r}*12",
         "N":f"=$J{r}*$E{r}",
         "O":f"=$N{r}*12",
         "P":f"=$J{r}*12",
         "Q":f'=IF($C{r}="EE",INDEX({PDI},MATCH($A{r},{PN},0)),INDEX({PDF},MATCH($A{r},{PN},0)))',
         "R":f'=IF($C{r}="EE",INDEX({POI},MATCH($A{r},{PN},0)),INDEX({POF},MATCH($A{r},{PN},0)))',
         "S":f"=$P{r}+$R{r}",
         "T":f'=$A{r}&"|"&$C{r}',
        }
        for col,formula in f.items():
            cell=c3[f"{col}{r}"]; cell.value=formula
            cell.font=GRN if col in ("D","F","G","H","Q","R") else INK
        for col in ("F","G","H","I","J","L","N"): c3[f"{col}{r}"].number_format=CUR2
        for col in ("M","O","P","Q","R","S"): c3[f"{col}{r}"].number_format=CUR
        c3[f"K{r}"].number_format=PCT
        c3[f"D{r}"].alignment=CTR; c3[f"E{r}"].alignment=CTR
        c3[f"S{r}"].font=Font(name=F,size=10,bold=True)
        for col in range(1,20): c3.cell(row=r,column=col).border=BOX
        if tcode=="EC":
            for col in range(1,20):
                c3.cell(row=r,column=col).fill=OFFFILL
                if c3.cell(row=r,column=col).font.color is None or c3.cell(row=r,column=col).font.color.rgb=="FF000000":
                    c3.cell(row=r,column=col).font=GREY
        elif tcode in ("ES","FAM"):
            for col in range(1,20): c3.cell(row=r,column=col).fill=BAND
        r+=1
    grp_last=r-1
    CT_TOTROWS.append(r)
    c3.cell(row=r,column=1,value=row[0]+"  -  TOTAL").font=BLD
    c3.cell(row=r,column=2,value="All tiers counted in totals").font=SM
    c=c3.cell(row=r,column=5,value=f"=SUM(E{grp_first}:E{grp_last})"); c.font=BLD; c.alignment=CTR
    for col,letter,fmt in ((12,"L",CUR2),(13,"M",CUR),(14,"N",CUR2),(15,"O",CUR)):
        c=c3.cell(row=r,column=col,value=f"=SUM({letter}{grp_first}:{letter}{grp_last})")
        c.font=BLD; c.number_format=fmt
    c=c3.cell(row=r,column=19,value=f"=SUMPRODUCT($S{grp_first}:$S{grp_last},$E{grp_first}:$E{grp_last})")
    c.font=BLD; c.number_format=CUR
    c3.cell(row=r,column=20,value="TOTAL ROW - not a lookup key").font=SM
    for col in range(1,20):
        c3.cell(row=r,column=col).border=BOX; c3.cell(row=r,column=col).fill=TOTFILL
    c3.row_dimensions[r].height=20
    r+=2
CT_LAST=CT_TOTROWS[-1]
c3.cell(row=CT_LAST+3,column=1,
  value=("HSA funding is deliberately not modeled. On the HDHP 3400/5000 an employee can pay part of that worst case with pretax "
         "dollars, so the real cost is lower - but by an amount only the employee controls.")).font=SUB
c3.cell(row=CT_LAST+4,column=1,
  value=("The 'all enrolled' and 'all in tier' columns multiply by how many people sit in that tier, so they sum to the whole spend "
         "IF every enrolled employee took that one plan. Actual elections are on the Elections tab.")).font=SUB
c3.cell(row=4,column=20,value="Row key").font=HDF
c3.cell(row=4,column=20).fill=HFILL; c3.cell(row=4,column=20).border=BOX
for col,w in [("A",24),("B",21),("C",7),("D",11),("E",10),("F",14),("G",16),("H",15),("I",15),("J",16),
              ("K",11),("L",17),("M",17),("N",17),("O",17),("P",16),("Q",13),("R",13),("S",18)]:
    c3.column_dimensions[col].width=w
c3.column_dimensions["T"].width=30; c3.column_dimensions["T"].hidden=True
c3.freeze_panes="D5"; c3.sheet_view.showGridLines=False
wb.defined_names.add(openpyxl.workbook.defined_name.DefinedName(
    "CostKey", attr_text=f"'Cost by Tier'!$T${CT_FIRST}:$T${CT_LAST}"))
CT="'Cost by Tier'!"
KEYM='MATCH($E{r}&"|"&$D{r},CostKey,0)'
print("cost by tier ok", CT_FIRST, CT_LAST)

# ================= 4. ELECTIONS =================
el=wb.create_sheet("Elections")
el["A1"]="Final elections - what each person actually chose"; el["A1"].font=TTL
el["A2"]=("This is the authoritative cost to Onyx. Pick each person's plan in the yellow column; the Combined tab's medical lines "
          "read the TOTAL row below.")
el["A2"].font=SUB
el["A3"]="EDIT THE YELLOW CELLS. Enrolling? and tier come from the Contributions tab."
el["A3"].font=Font(name=F,size=10,bold=True,color="8E2F2A")
sec(el,5,"FINAL PLAN SELECTION BY PERSON",13)
hdr(el,6,["Employee","Enrolling?","Tier","Code","ELECTED PLAN","In totals?","Monthly premium",
          "Onyx $/mo","Employee $/mo","Total $/mo","Onyx $/yr","Employee $/yr","Total $/yr","Worst case $/yr (each)"],h=40)
EL_FIRST=7
for i,(n,pay,tier,code,yn,_note) in enumerate(EMP):
    r=EL_FIRST+i; m=KEYM.format(r=r)
    gate=f'IF(OR($B{r}<>"Yes",$F{r}<>"Yes"),0,'
    el.cell(row=r,column=1,value=n).font=INK
    c=el.cell(row=r,column=2,value=f"=Contributions!$E${EMP_FIRST+i}"); c.font=GRN; c.alignment=CTR
    el.cell(row=r,column=3,value=tier).font=INK
    c=el.cell(row=r,column=4,value=f"=Contributions!$D${EMP_FIRST+i}"); c.font=GRN; c.alignment=CTR
    c=el.cell(row=r,column=5,value=ANGLE[0][0]); c.font=BLUB; c.fill=YEL
    c=el.cell(row=r,column=6,value=f"=INDEX({INCRNG},MATCH($D{r},{CODERNG},0))"); c.font=GRN; c.alignment=CTR
    c=el.cell(row=r,column=7,value=f"=INDEX({CT}$F${CT_FIRST}:$F${CT_LAST},{m})"); c.font=GRN; c.number_format=CUR2
    c=el.cell(row=r,column=8,value=f"={gate}INDEX({CT}$I${CT_FIRST}:$I${CT_LAST},{m}))"); c.font=INK; c.number_format=CUR2
    c=el.cell(row=r,column=9,value=f"={gate}INDEX({CT}$J${CT_FIRST}:$J${CT_LAST},{m}))"); c.font=INK; c.number_format=CUR2
    c=el.cell(row=r,column=10,value=f"=$H{r}+$I{r}"); c.font=BLD; c.number_format=CUR2
    c=el.cell(row=r,column=11,value=f"=$H{r}*12"); c.font=INK; c.number_format=CUR
    c=el.cell(row=r,column=12,value=f"=$I{r}*12"); c.font=INK; c.number_format=CUR
    c=el.cell(row=r,column=13,value=f"=$J{r}*12"); c.font=BLD; c.number_format=CUR
    c=el.cell(row=r,column=14,value=f"={gate}INDEX({CT}$S${CT_FIRST}:$S${CT_LAST},{m}))")
    c.font=INK; c.number_format=CUR
    el.row_dimensions[r].height=22
    for col in range(1,15): el.cell(row=r,column=col).border=BOX
    if yn!="Yes":
        for col in range(1,15): el.cell(row=r,column=col).fill=BAND
        el.cell(row=r,column=5).fill=YEL
EL_LAST=EL_FIRST+len(EMP)-1
dvp=DataValidation(type="list",formula1='"{}"'.format(",".join(x[0] for x in ANGLE)),
                   allow_blank=False,showDropDown=False)
el.add_data_validation(dvp); dvp.add(f"E{EL_FIRST}:E{EL_LAST}")
EL_TOT=EL_LAST+1
el.cell(row=EL_TOT,column=1,value="TOTAL - ONYX AND EMPLOYEES").font=BLD
c=el.cell(row=EL_TOT,column=2,value=f'=COUNTIF(B{EL_FIRST}:B{EL_LAST},"Yes")&" enrolled"'); c.font=BLD; c.alignment=CTR
for col,letter,fmt in ((8,"H",CUR2),(9,"I",CUR2),(10,"J",CUR2),(11,"K",CUR),(12,"L",CUR),(13,"M",CUR),(14,"N",CUR)):
    c=el.cell(row=EL_TOT,column=col,value=f"=SUM({letter}{EL_FIRST}:{letter}{EL_LAST})")
    c.font=BLD; c.number_format=fmt
for col in range(1,15):
    el.cell(row=EL_TOT,column=col).border=BOX; el.cell(row=EL_TOT,column=col).fill=TOTFILL
el.cell(row=EL_TOT+2,column=1,value="ONYX - MONTHLY").font=BLD
c=el.cell(row=EL_TOT+2,column=3,value=f"=$H${EL_TOT}"); c.font=BLD; c.number_format=CUR2
el.cell(row=EL_TOT+3,column=1,value="ONYX - ANNUAL").font=BLD
c=el.cell(row=EL_TOT+3,column=3,value=f"=$K${EL_TOT}"); c.font=BLD; c.number_format=CUR
el.cell(row=EL_TOT+4,column=1,value="Employees - monthly").font=INK
c=el.cell(row=EL_TOT+4,column=3,value=f"=$I${EL_TOT}"); c.font=INK; c.number_format=CUR2
el.cell(row=EL_TOT+5,column=1,value="Employees - annual").font=INK
c=el.cell(row=EL_TOT+5,column=3,value=f"=$L${EL_TOT}"); c.font=INK; c.number_format=CUR
for rr in range(EL_TOT+2,EL_TOT+6):
    for col in range(1,4): el.cell(row=rr,column=col).border=BOX
for rr in (EL_TOT+2,EL_TOT+3):
    for col in range(1,4): el.cell(row=rr,column=col).fill=TOTFILL
el.cell(row=EL_TOT+7,column=1,value=(
  "A person whose tier is excluded on the Contributions tab, or who is not enrolling, shows $0 on both sides - their premium still "
  "displays so the record is complete.")).font=SUB
el.cell(row=EL_TOT+8,column=1,value=(
  "Everyone defaults to the HDHP 3400/5000. Change the ELECTED PLAN cells as people choose; nothing else needs touching.")).font=SUB
el.cell(row=EL_TOT+9,column=1,value=(
  "Worst case = that person's annual premium share plus their household's in-network out-of-pocket maximum, on the plan they elected. "
  "Premiums never count toward an out-of-pocket maximum, so the two add without overlapping. The Combined tab reads the TOTAL of "
  "this column.")).font=SUB

# --- saved scenarios: name a column, type a percentage per person, read the totals ---
SC_NAME=23+2; SC_SUB=SC_NAME+1; SC_FIRST=SC_SUB+1
SC_LAST=SC_FIRST+len(EMP)-1
SC_ONM=SC_LAST+1; SC_ONY=SC_ONM+1; SC_EEM=SC_ONY+1; SC_EEY=SC_EEM+1
SC_DO=SC_EEY+1; SC_DE=SC_DO+1
PAIRS=[("E","F"),("G","H"),("I","J"),("K","L")]
SCEN=[("1.  Current - 50% all",   {}),
      ("2.  Christine at 55%",    {"Christine Johnson":0.55}),
      ("3.  (name this one)",     {}),
      ("4.  (name this one)",     {})]
PREMTOT=(f'SUMPRODUCT(($B${EL_FIRST}:$B${EL_LAST}="Yes")*($F${EL_FIRST}:$F${EL_LAST}="Yes")'
         f'*$G${EL_FIRST}:$G${EL_LAST})')

sec(el,23,"SAVED SCENARIOS   <- name a column, type a percentage for each person, read the totals",12)
el.cell(row=24,column=1,value=(
  "A scratchpad. Nothing here feeds the Combined tab - it is only for comparing. Everyone's plan, tier and enrolment come from the "
  "table above, so the only thing that changes between columns is the percentage Onyx funds.")).font=SUB
el.cell(row=SC_NAME,column=1,value="SCENARIO").font=HDF
el.cell(row=SC_NAME,column=1).fill=HFILL; el.cell(row=SC_NAME,column=1).border=BOX
el.cell(row=SC_NAME,column=1).alignment=CTR
for ci,(pc,dc) in enumerate(PAIRS):
    el.merge_cells(f"{pc}{SC_NAME}:{dc}{SC_NAME}")
    c=el[f"{pc}{SC_NAME}"]; c.value=SCEN[ci][0]; c.font=BLUB; c.fill=YEL
    c.alignment=Alignment(horizontal="center",vertical="center")
    for col in (pc,dc): el[f"{col}{SC_NAME}"].border=BOX
el.row_dimensions[SC_NAME].height=22
hdr(el,SC_SUB,["Employee","Code","Elected plan","Premium $/mo"],h=32)
for pc,dc in PAIRS:
    for col,lab in ((pc,"Onyx %"),(dc,"Onyx $/mo")):
        c=el[f"{col}{SC_SUB}"]; c.value=lab; c.font=HDF; c.fill=HFILL; c.border=BOX
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
for i,(n,pay,tier,code,yn,_note) in enumerate(EMP):
    r=SC_FIRST+i; u=EL_FIRST+i
    gate=f'IF(OR($B${u}<>"Yes",$F${u}<>"Yes"),0,'
    c=el.cell(row=r,column=1,value=f"=$A${u}"); c.font=GRN
    c=el.cell(row=r,column=2,value=f"=$D${u}"); c.font=GRN; c.alignment=CTR
    c=el.cell(row=r,column=3,value=f"=$E${u}"); c.font=GRN
    c=el.cell(row=r,column=4,value=f"=$G${u}"); c.font=GRN; c.number_format=CUR2
    for ci,(pc,dc) in enumerate(PAIRS):
        c=el[f"{pc}{r}"]; c.value=SCEN[ci][1].get(n,PCT_DEFAULT)
        c.font=BLUB; c.fill=YEL; c.number_format=PCT
        c=el[f"{dc}{r}"]
        c.value=f"={gate}MIN($D{r},MAX({pc}{r}*$D{r},{MINCELL})))"
        c.font=INK; c.number_format=CUR2
    for col in range(1,13): el.cell(row=r,column=col).border=BOX
    if yn!="Yes":
        for col in range(1,13): el.cell(row=r,column=col).fill=BAND
        for pc,_dc in PAIRS: el[f"{pc}{r}"].fill=YEL
for lab,rr,bold in (("ONYX - MONTHLY",SC_ONM,True),("ONYX - ANNUAL",SC_ONY,True),
                    ("Employees - monthly",SC_EEM,False),("Employees - annual",SC_EEY,False),
                    ("Change to ONYX $/yr vs Scenario 1",SC_DO,True),
                    ("Change to employees $/yr vs Scenario 1",SC_DE,False)):
    el.cell(row=rr,column=1,value=lab).font=BLD if bold else INK
    for pc,dc in PAIRS:
        c=el[f"{dc}{rr}"]
        if rr==SC_ONM: c.value=f"=SUM({dc}{SC_FIRST}:{dc}{SC_LAST})"; c.number_format=CUR2
        elif rr==SC_ONY: c.value=f"={dc}{SC_ONM}*12"; c.number_format=CUR
        elif rr==SC_EEM: c.value=f"={PREMTOT}-{dc}{SC_ONM}"; c.number_format=CUR2
        elif rr==SC_EEY: c.value=f"={dc}{SC_EEM}*12"; c.number_format=CUR
        elif rr==SC_DO: c.value=f"={dc}{SC_ONY}-$F${SC_ONY}"; c.number_format=CUR
        else: c.value=f"={dc}{SC_EEY}-$F${SC_EEY}"; c.number_format=CUR
        c.font=BLD if bold else INK
    for col in range(1,13): el.cell(row=rr,column=col).border=BOX
    if rr in (SC_ONM,SC_ONY,SC_DO):
        for col in range(1,13): el.cell(row=rr,column=col).fill=TOTFILL
el.cell(row=SC_DE+2,column=1,value=(
  f'=IF($H${SC_DO}=0,"Scenario 2 costs Onyx exactly what Scenario 1 does.",'
  f'"Scenario 2 costs Onyx "&TEXT(ABS($H${SC_DO}),"$#,##0")&IF($H${SC_DO}>0," MORE"," LESS")&'
  f'" a year than Scenario 1 - and moves that same amount "&IF($H${SC_DO}>0,"off","onto")&" the employees.")')
  ).font=Font(name=F,size=11,bold=True,color="1D5B72")
el.cell(row=SC_DE+3,column=1,value=(
  "Onyx never pays less than Angle's minimum funding, so a percentage low enough to fall under it gets topped up automatically - "
  "which is why the totals stop falling below a floor.")).font=SUB
el.cell(row=SC_DE+4,column=1,value=(
  "Scenario 1 is the baseline the two change rows measure against. Overwrite any column's name and percentages to try something else.")).font=SUB

for col,w in [("A",30),("B",13),("C",21),("D",13),("E",24),("F",11),("G",16),("H",14),("I",16),("J",14),("K",14),("L",16),("M",14),("N",20)]:
    el.column_dimensions[col].width=w
el.freeze_panes="F7"; el.sheet_view.showGridLines=False
print("elections ok", EL_FIRST, EL_LAST, EL_TOT)

# ================= 5. BY EMPLOYEE =================
e=wb.create_sheet("By Employee")
e["A1"]="Every Angle plan, priced for each person"; e["A1"].font=TTL
e["A2"]="What each plan would cost each person. The plan actually chosen is on the Elections tab; this grid is the menu behind it."
e["A2"].font=SUB
hdr(e,4,["Employee","Enrolling?","Tier","Code","Plan","In totals?","Monthly premium","Onyx pays $/mo",
         "Employee $/mo","Employee $/yr","Deductible","OOP max","Worst case $/yr"],h=34)
r=5; E_FIRST=5
for ei,(n,pay,tier,code,yn,_note) in enumerate(EMP):
    for row in ANGLE:
        m=f'MATCH($E{r}&"|"&$D{r},CostKey,0)'
        e.cell(row=r,column=1,value=n).font=INK
        e[f"B{r}"]=f"=Contributions!$E${EMP_FIRST+ei}"; e[f"B{r}"].font=GRN; e[f"B{r}"].alignment=CTR
        e.cell(row=r,column=3,value=tier).font=INK
        e.cell(row=r,column=4,value=code).font=INK; e.cell(row=r,column=4).alignment=CTR
        e.cell(row=r,column=5,value=row[0]).font=INK
        for col,src in [("F","D"),("G","F"),("H","I"),("I","J"),("J","P"),("K","Q"),("L","R"),("M","S")]:
            e[f"{col}{r}"]=f"=INDEX({CT}${src}${CT_FIRST}:${src}${CT_LAST},{m})"
            e[f"{col}{r}"].font=GRN
        e[f"F{r}"].alignment=CTR
        for col in ("G","H","I"): e[f"{col}{r}"].number_format=CUR2
        for col in ("J","K","L","M"): e[f"{col}{r}"].number_format=CUR
        e[f"M{r}"].font=Font(name=F,size=10,bold=True,color="008000")
        for col in range(1,14): e.cell(row=r,column=col).border=BOX
        if yn!="Yes":
            for col in range(1,14): e.cell(row=r,column=col).fill=BAND
        r+=1
E_LAST=r-1
for col,w in [("A",22),("B",11),("C",21),("D",7),("E",24),("F",11),("G",14),("H",14),("I",14),("J",14),("K",13),("L",13),("M",16)]:
    e.column_dimensions[col].width=w
e.freeze_panes="G5"; e.sheet_view.showGridLines=False
e.auto_filter.ref=f"A4:M{E_LAST}"
print("by employee ok", E_FIRST, E_LAST)

# ================= 6. QUESTCO =================
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
for i,(line,basis,rate,el2,note) in enumerate(QR):
    r=6+i
    q.cell(row=r,column=1,value=line).font=INK
    q.cell(row=r,column=2,value=basis).font=SM
    c=q.cell(row=r,column=3,value=rate); c.font=BLUB; c.fill=YEL; c.number_format=CUR2
    c=q.cell(row=r,column=5,value=el2); c.alignment=CTR
    c.font=SM if el2=="Required" else BLUB
    if el2!="Required": c.fill=YEL
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
for i,(b,el2,design,mon,note) in enumerate(ANC):
    r=15+i
    q.cell(row=r,column=1,value=b).font=INK
    c=q.cell(row=r,column=2,value=el2); c.font=BLUB; c.fill=YEL; c.alignment=CTR
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
     ("401(k)","Large-group multiple employer plan, $3.00 per participating employee per quarter, no employer cost unless Onyx matches. Confirmed it does not require annual re-signature by Onyx.","High. See the 401(k) tab - the match is the only part Onyx controls."),
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

# ================= 7. 401(k) =================
k=wb.create_sheet("401(k)")
k["A1"]="401(k) through Questco - match modelling"; k["A1"].font=TTL
k["A2"]=("Questco charges $3.00 per participating employee per quarter and nothing to Onyx unless Onyx chooses to match. "
         "A match is NOT required - see block 4. Everything below is a lever.")
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
       "Questco's quoted fee - $12 a year per participant.")
setrow(8,"Fee paid by","Participant",None,
       "Set to Employer to have Onyx absorb it. Questco's statement is that Onyx pays nothing unless it matches, so Participant is the default - worth confirming.",f1=None)
setrow(9,"IRS compensation limit",350000,None,
       "CONFIRM the current-year 401(a)(17) figure. Nobody here is close, so it does not bind.",f1=CUR)
setrow(10,"IRS elective deferral limit",24000,None,
       "CONFIRM the current-year 402(g) figure. Not binding at the deferral rates below.",f1=CUR)
dvf=DataValidation(type="list",formula1='"Participant,Employer"',allow_blank=False,showDropDown=False)
k.add_data_validation(dvf); dvf.add("C8")
setrow(12,"Match type","None",None,
       "None, Tiered match, or Non-elective. Defaults to None because nothing has been committed. Block 3 prices all three at once.",f1=None)
setrow(13,"Tier 1 - match rate",1.00,0.03,
       "Match 100% of the first 3% of pay deferred. With Tier 2 this is the safe harbor basic formula.",f1=PCT,f2=PCT,yellow2=True)
setrow(14,"Tier 2 - match rate",0.50,0.02,
       "Match 50% of the next 2% of pay. Anyone deferring 5% or more earns the full 4% of pay.",f1=PCT,f2=PCT,yellow2=True)
setrow(15,"Non-elective rate",0.03,None,
       "Used only when Match type is Non-elective. 3% of pay to every participating employee, whether or not they defer.",f1=PCT)
dvm=DataValidation(type="list",formula1='"None,Tiered match,Non-elective"',allow_blank=False,showDropDown=False)
k.add_data_validation(dvm); dvm.add("C12")

sec(k,17,"2.  WHO PARTICIPATES AND WHAT THEY DEFER",10)
hdr(k,18,["Employee","Eligible pay","Participating?","Deferral % of pay","Employee deferral $/yr",
          "Onyx match $/yr (selected type)","Match as % of pay","Participant fee $/yr",
          "IF safe harbor match","IF 3% non-elective"],h=40)
K_FIRST=19
SHM='MIN($B{r},$C$9)*($C$13*MIN($D{r},$D$13)+$C$14*MIN(MAX($D{r}-$D$13,0),$D$14))'
NEL='MIN($B{r},$C$9)*$C$15'
for i,(n,pay,tier,code,yn,_note) in enumerate(EMP):
    r=K_FIRST+i
    k.cell(row=r,column=1,value=n).font=INK
    c=k.cell(row=r,column=2,value=pay); c.font=BLU; c.number_format=CUR
    c=k.cell(row=r,column=3,value="Yes"); c.font=BLUB; c.fill=YEL; c.alignment=CTR
    c=k.cell(row=r,column=4,value=0.06); c.font=BLUB; c.fill=YEL; c.number_format=PCT
    c=k.cell(row=r,column=5,value=f'=IF($C{r}<>"Yes",0,MIN(MIN($B{r},$C$9)*$D{r},$C$10))')
    c.font=INK; c.number_format=CUR
    c=k.cell(row=r,column=6,value=(
        f'=IF($C{r}<>"Yes",0,IF($C$12="None",0,'
        f'IF($C$12="Non-elective",$J{r},$I{r})))'))
    c.font=INK; c.number_format=CUR
    c=k.cell(row=r,column=7,value=f"=IF($B{r}=0,0,$F{r}/$B{r})"); c.font=INK; c.number_format=PCT
    c=k.cell(row=r,column=8,value=f'=IF($C{r}<>"Yes",0,$C$7*4)'); c.font=INK; c.number_format=CUR
    c=k.cell(row=r,column=9,value=f'=IF($C{r}<>"Yes",0,{SHM.format(r=r)})'); c.font=SM; c.number_format=CUR
    c=k.cell(row=r,column=10,value=f'=IF($C{r}<>"Yes",0,{NEL.format(r=r)})'); c.font=SM; c.number_format=CUR
    for col in range(1,11): k.cell(row=r,column=col).border=BOX
K_LAST=K_FIRST+len(EMP)-1
dvpk=DataValidation(type="list",formula1='"Yes,No"',allow_blank=False,showDropDown=False)
k.add_data_validation(dvpk); dvpk.add(f"C{K_FIRST}:C{K_LAST}")
r=K_LAST+1; K_TOT=r
k.cell(row=r,column=1,value="TOTAL").font=BLD
k.cell(row=r,column=2,value=f"=SUM(B{K_FIRST}:B{K_LAST})").font=BLD
k.cell(row=r,column=2).number_format=CUR
k.cell(row=r,column=3,value=f'=COUNTIF(C{K_FIRST}:C{K_LAST},"Yes")&" in"').font=BLD
k.cell(row=r,column=3).alignment=CTR
for col,letter in ((5,"E"),(6,"F"),(8,"H"),(9,"I"),(10,"J")):
    c=k.cell(row=r,column=col,value=f"=SUM({letter}{K_FIRST}:{letter}{K_LAST})")
    c.font=BLD; c.number_format=CUR
c=k.cell(row=r,column=7,value=f"=IF(B{r}=0,0,F{r}/B{r})"); c.font=BLD; c.number_format=PCT
for col in range(1,11):
    k.cell(row=r,column=col).border=BOX; k.cell(row=r,column=col).fill=TOTFILL
k.cell(row=K_TOT+1,column=1,
  value="Deferral rates default to 6%, which fully earns the safe harbor basic match. Nobody has elected anything yet - this is an assumption, and it is what drives the two right-hand columns.").font=SUB

sec(k,K_TOT+3,"3.  THE THREE OPTIONS, PRICED SIDE BY SIDE",6)
hdr(k,K_TOT+4,["Option","Onyx cost $/yr","Onyx cost $/mo","What it buys"])
OPT=[("No employer contribution",0,"=B{r}/12",
      "Cheapest. But the plan is then subject to nondiscrimination testing every year, and once it becomes top-heavy Onyx owes the minimum anyway."),
     ("Safe harbor basic match (100% of first 3%, 50% of next 2%)",f"=I{K_TOT}","=B{r}/12",
      "Exempts the plan from nondiscrimination testing and, on its own, from the top-heavy minimum. Only the people who defer receive it."),
     ("Safe harbor 3% non-elective",f"=J{K_TOT}","=B{r}/12",
      "Same exemptions. Goes to every participating employee whether or not they defer, so it costs Onyx money for non-savers."),
    ]
for i,(opt,ann,mon,note) in enumerate(OPT):
    r=K_TOT+5+i
    k.cell(row=r,column=1,value=opt).font=BLD if i==0 else INK
    c=k.cell(row=r,column=2,value=ann); c.font=BLD; c.number_format=CUR
    c=k.cell(row=r,column=3,value=mon.format(r=r)); c.font=INK; c.number_format=CUR2
    k.cell(row=r,column=4,value=note).font=SM; k.cell(row=r,column=4).alignment=WRAP
    k.row_dimensions[r].height=40
    for col in range(1,5): k.cell(row=r,column=col).border=BOX
K_OPT=K_TOT+5

sec(k,K_TOT+9,"4.  WHAT ONYX IS ACTUALLY COSTED AT, GIVEN THE MATCH TYPE ABOVE",5)
hdr(k,K_TOT+10,["Line","Monthly","Annual","Note"])
KS=[("Employer match / non-elective",f"=F{K_TOT}",
     "Deductible to Onyx. Unlike health premiums, a contribution for Steven as a >2% S-corp shareholder is a normal deductible plan contribution, not W-2 wages."),
    ("Participant fees absorbed by Onyx",f'=IF($C$8="Employer",H{K_TOT},0)',
     "Zero while the fee sits on participants."),
   ]
for i,(line,ann,note) in enumerate(KS):
    r=K_TOT+11+i
    k.cell(row=r,column=1,value=line).font=INK
    c=k.cell(row=r,column=2,value=f"=C{r}/12"); c.font=INK; c.number_format=CUR2
    c=k.cell(row=r,column=3,value=ann); c.font=INK; c.number_format=CUR
    k.cell(row=r,column=4,value=note).font=SM; k.cell(row=r,column=4).alignment=WRAP
    k.row_dimensions[r].height=34
    for col in range(1,5): k.cell(row=r,column=col).border=BOX
r=K_TOT+13; K_COST=r
k.cell(row=r,column=1,value="ONYX 401(k) COST").font=BLD
c=k.cell(row=r,column=2,value=f"=SUM(B{K_TOT+11}:B{K_TOT+12})"); c.font=BLD; c.number_format=CUR2
c=k.cell(row=r,column=3,value=f"=SUM(C{K_TOT+11}:C{K_TOT+12})"); c.font=BLD; c.number_format=CUR
for col in range(1,5):
    k.cell(row=r,column=col).border=BOX; k.cell(row=r,column=col).fill=TOTFILL

sec(k,K_COST+2,"5.  IS A MATCH REQUIRED?   No - but read this before choosing None",5)
for i,t in enumerate([
  "NO LAW REQUIRES AN EMPLOYER MATCH. A 401(k) can be deferral-only, funded entirely by employees. Questco's fee does not change either way.",
  "Two testing rules can still force money in. Nondiscrimination testing compares what the highly compensated defer against what everyone else defers; if the gap is too wide, the highly compensated get deferrals refunded. Steven is highly compensated by ownership and Lisa very likely by pay, so two of five would sit on that side of the test. (The IRS calls these the ADP and ACP tests - an actuarial term that has nothing to do with the payroll company of the same initials.)",
  "Top-heavy: when key employees hold more than 60% of plan assets, the employer owes a minimum contribution of up to 3% of pay to every non-key employee. With Steven owning 100% this plan is likely to go top-heavy - though usually not in year one, since the test looks at the prior year end and a new plan starts at zero.",
  "A safe harbor design - the basic match or the 3% non-elective - buys exemption from nondiscrimination testing, and from the top-heavy minimum when the safe harbor contribution is the only employer money in the plan. That is the real trade: a known, capped cost instead of an uncertain one.",
  "So the practical choice is not match or no match. It is: accept testing and a probable top-heavy bill later, or take a safe harbor now and know the number. Block 3 prices both safe harbor routes.",
  "Confirm the current-year compensation, deferral and highly-compensated thresholds, and the top-heavy position, with the plan's third-party administrator before adopting.",
  "Ask Questco whether an adopting employer in their multiple employer plan still claims the SECURE 2.0 start-up credits - up to $5,000/yr for three years plus $500/yr for auto-enrolment. Still unanswered.",
]):
    c=k.cell(row=K_COST+3+i,column=1,value=u"•  "+t); c.font=INK; c.alignment=WRAP
    k.row_dimensions[K_COST+3+i].height=32
for col,w in [("A",58),("B",16),("C",18),("D",62),("E",62),("F",18),("G",15),("H",15),("I",18),("J",18)]:
    k.column_dimensions[col].width=w
k.sheet_view.showGridLines=False
print("401k ok  rows",K_FIRST,K_LAST,"| TOT",K_TOT,"| OPT",K_OPT,"| cost",K_COST)

# ================= 8. COMBINED =================
CB=wb.create_sheet("Combined")
PLAN_A=f"{CT}$A${CT_FIRST}:$A${CT_LAST}"
ONYX_Y=f"{CT}$M${CT_FIRST}:$M${CT_LAST}"
EMP_Y=f"{CT}$O${CT_FIRST}:$O${CT_LAST}"
WORST_O=f"{CT}$S${CT_FIRST}:$S${CT_LAST}"
ENR_D=f"{CT}$E${CT_FIRST}:$E${CT_LAST}"
QO="Questco!"; KO="'401(k)'!"; EO="Elections!"
NONMED_A=f"{QO}$G$6+{QO}$G$7+{QO}$G$8+{QO}$E$18+{KO}$C${K_COST}"

CB["A1"]="Angle + Questco - monthly and annual projection"; CB["A1"].font=TTL
CB["A2"]="Block 1 reads the actual elections on the Elections tab. Block 3 shows what each plan would cost if every enrolled employee took it."
CB["A2"].font=SUB

sec(CB,5,"1.  WHAT ONYX PAYS   (at the elections currently entered)",6)
hdr(CB,6,["Cost line","Vendor","Monthly","Annual","Note"])
LINES=[("Medical - employer contribution","Angle / CapFi",
        f"={EO}$H${EL_TOT}",f"={EO}$K${EL_TOT}",
        "From the Elections tab. On the percentage basis this moves with the plan each person elects."),
       ("Administrative fee","Questco",f"={QO}F6",f"={QO}G6","$117 per employee per month."),
       ("Workers' comp + EPLI","Questco",f"={QO}F7",f"={QO}G7","Subject to underwriting."),
       ("Cyber liability","Questco",f"={QO}F8",f"={QO}G8",
        "Zero while it is set to No on the Questco tab. Onyx is placing this separately."),
       ("Employer-paid ancillary","Questco",f"={QO}D18",f"={QO}E18",
        "Life, LTD and STD as elected on the Questco tab."),
       ("401(k) - employer cost","Questco",f"={KO}B{K_COST}",f"={KO}C{K_COST}",
        "Match, plus participant fees only if Onyx absorbs them. Match type defaults to None."),
      ]
for i,(line,ven,mon,ann,note) in enumerate(LINES):
    r=7+i
    CB.cell(row=r,column=1,value=line).font=INK
    CB.cell(row=r,column=2,value=ven).font=SM
    c=CB.cell(row=r,column=3,value=mon); c.font=INK; c.number_format=CUR2
    c=CB.cell(row=r,column=4,value=ann); c.font=INK; c.number_format=CUR
    CB.cell(row=r,column=5,value=note).font=SM; CB.cell(row=r,column=5).alignment=WRAP
    CB.row_dimensions[r].height=28
    for col in range(1,6): CB.cell(row=r,column=col).border=BOX
CB["A13"]="ONYX RECURRING TOTAL"; CB["A13"].font=BLD
CB["C13"]="=SUM(C7:C12)"; CB["C13"].font=BLD; CB["C13"].number_format=CUR2
CB["D13"]="=SUM(D7:D12)"; CB["D13"].font=BLD; CB["D13"].number_format=CUR
CB["A14"]="Implementation (first year only)"; CB["A14"].font=INK
CB["B14"]="Questco"; CB["B14"].font=SM
CB["D14"]=f"={QO}G10"; CB["D14"].font=INK; CB["D14"].number_format=CUR
CB["A15"]="ONYX FIRST-YEAR TOTAL"; CB["A15"].font=BLD
CB["C15"]="=D15/12"; CB["C15"].font=BLD; CB["C15"].number_format=CUR2
CB["D15"]="=D13+D14"; CB["D15"].font=BLD; CB["D15"].number_format=CUR
for r2 in (13,14,15):
    for col in range(1,6): CB.cell(row=r2,column=col).border=BOX
for r2 in (13,15):
    for col in range(1,6): CB.cell(row=r2,column=col).fill=TOTFILL
CB["A16"]=("The 401(k) line is a MODELLED match, not a committed cost. Match type is set to None on the 401(k) tab, so it reads zero. "
           "Block 3 of that tab prices the two safe harbor alternatives.")
CB["A16"].font=Font(name=F,size=9,bold=True,italic=True,color="8E2F2A")
CB["A17"]=("Payroll taxes of about $52,176 a year are NOT in this total. Onyx owes FICA, FUTA and SUTA whether or not there is a PEO, "
           "so they are a pass-through on the Questco invoice rather than a cost of the relationship. They are itemised on the Questco tab.")
CB["A17"].font=SUB
CB["A18"]="Employee + child(ren) is excluded from every figure on this tab. Its rates are kept on the Angle Plans and Cost by Tier tabs for the record."
CB["A18"].font=GREYI

sec(CB,20,"2.  WHAT THE EMPLOYEES PAY, AND THE COMBINED PICTURE",6)
hdr(CB,21,["Line","Monthly","Annual","Note"])
EL2=[("Employee medical deductions, all enrolled",f"={EO}$I${EL_TOT}",f"={EO}$L${EL_TOT}",
      "From the Elections tab. Premium only. Pretax for everyone except Steven."),
     ("Employee 401(k) deferrals",f"={KO}E{K_TOT}/12",f"={KO}E{K_TOT}",
      "Their own savings, not a cost - shown so the payroll picture is complete."),
     ("Employee 401(k) fees",f"=IF({KO}$C$8=\"Employer\",0,{KO}H{K_TOT}/12)",
      f"=IF({KO}$C$8=\"Employer\",0,{KO}H{K_TOT})","$3 per quarter each while the fee sits on participants."),
     ("Onyx + employees, medical premium only",f"={EO}$J${EL_TOT}",f"={EO}$M${EL_TOT}",
      "The true cost of the Angle plan itself, both sides."),
     ("Onyx + employees, all in","=C13+B22+B24","=D13+C22+C24",
      "Adds the Questco fee, elected ancillary, the 401(k) and the participant fees."),
     ("Combined worst case, one bad year","",
      f"=D13+{EO}$N${EL_TOT}",
      "Each enrolled household on the plan it actually elected, all hitting their out-of-pocket maximum in the same year. A ceiling, not a forecast."),
    ]
for i,(line,mon,ann,note) in enumerate(EL2):
    r=22+i
    CB.cell(row=r,column=1,value=line).font=BLD if i>=3 else INK
    if mon:
        c=CB.cell(row=r,column=2,value=mon); c.font=BLD if i>=3 else INK; c.number_format=CUR2
    c=CB.cell(row=r,column=3,value=ann); c.font=BLD if i>=3 else INK; c.number_format=CUR
    CB.cell(row=r,column=4,value=note).font=SM; CB.cell(row=r,column=4).alignment=WRAP
    CB.row_dimensions[r].height=28
    for col in range(1,5): CB.cell(row=r,column=col).border=BOX

sec(CB,29,"3.  ALL THREE PLANS SIDE BY SIDE   (annual, if every enrolled employee took that one plan)",7)
hdr(CB,30,["Angle plan","HSA","Onyx - medical","Onyx - all in","Employee medical","Onyx + employees","Combined worst case"])
for i,row in enumerate(ANGLE):
    r=31+i
    CB.cell(row=r,column=1,value=row[0]).font=INK
    CB.cell(row=r,column=2,value=row[1]).font=INK; CB.cell(row=r,column=2).alignment=CTR
    CB[f"C{r}"]=f"=SUMIFS({ONYX_Y},{PLAN_A},$A{r})"
    CB[f"D{r}"]=f"=C{r}+{NONMED_A}"
    CB[f"E{r}"]=f"=SUMIFS({EMP_Y},{PLAN_A},$A{r})"
    CB[f"F{r}"]=f'=D{r}+E{r}+IF({KO}$C$8="Employer",0,{KO}H{K_TOT})'
    CB[f"G{r}"]=f"=D{r}+SUMPRODUCT(({PLAN_A}=$A{r})*{WORST_O}*{ENR_D})"
    for col in "CDEFG":
        CB[f"{col}{r}"].font=INK; CB[f"{col}{r}"].number_format=CUR
    CB[f"D{r}"].font=BLD
    for col in range(1,8): CB.cell(row=r,column=col).border=BOX
    if i%2:
        for col in range(1,8): CB.cell(row=r,column=col).fill=BAND
CB["A35"]=("On the percentage basis Onyx's medical cost RISES with the plan, because Onyx pays 50% of whatever the premium is. "
           "Switch the basis to Fixed dollars on the Contributions tab and the column goes flat instead - Onyx pays the same on every "
           "plan and the employee absorbs the whole difference.")
CB["A35"].font=SUB
CB["A36"]="Angle offers no dental or vision. Questco quotes both as employee-paid, so neither appears in Onyx's cost here."
CB["A36"].font=SUB
for col,w in [("A",42),("B",18),("C",18),("D",18),("E",18),("F",20),("G",22)]:
    CB.column_dimensions[col].width=w
CB.sheet_view.showGridLines=False
print("combined ok")

wb.save("Onyx_Angle_Questco_Cost_Model.xlsx")
print("SAVED")
