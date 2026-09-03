# vendor, plan, network, hsa, ded_i, ded_f, oop_i, oop_f, plan_pays, pcp, spec, rx, er, EE, ES, EC, FAM
PLANS = [
# ---------------- ADP TotalSource / Aetna  (11 plans, deck pp.6-8) ----------------
("ADP","EPO HDHP 2000-100Copay TIF","Aetna EPO - no out-of-network","Y",2000,4000,3500,7000,1.00,"ded then $30","ded then $60","$10d / $45d / $70d","ded + $350",640.66,1377.42,1249.28,1986.04),
("ADP","EPO HDHP 2500-100Copay TIF","Aetna EPO - no out-of-network","Y",2500,5000,3500,7000,1.00,"ded then $30","ded then $60","$10d / $45d / $70d","ded + $350",606.87,1304.77,1183.39,1881.29),
("ADP","NATIONAL OA MC 500-80%","Aetna National OA Managed Choice - OON covered","N",500,1000,3500,7000,0.80,"$25","$50","$10 / $45 / $70","$350",882.91,1898.25,1721.67,2737.01),
("ADP","NATIONAL OA MC 1000-80%","Aetna National OA Managed Choice - OON covered","N",1000,2000,4500,9000,0.80,"$25","$50","$10 / $45 / $70","$500",768.21,1651.66,1498.01,2381.46),
("ADP","OA EPO HDHP 3500-80%","Aetna OA EPO - no out-of-network","Y",3500,7000,6350,12700,0.80,"20% ded","20% ded","$10d / $45d / $70d","ded + 20%",490.12,1053.75,955.73,1519.36),
("ADP","OA EPO 500-80%","Aetna OA EPO - no out-of-network","N",500,1000,3500,7000,0.80,"$25","$50","$10 / $45 / $70","$350",825.23,1774.24,1609.19,2558.21),
("ADP","OA EPO 1000-80%","Aetna OA EPO - no out-of-network","N",1000,2000,4500,9000,0.80,"$25","$50","$10 / $45 / $70","$500",749.15,1610.67,1460.84,2322.36),
("ADP","OA EPO 1500-100%","Aetna OA EPO - no out-of-network","N",1500,3000,5000,10000,1.00,"$30","$60","$10 / $45 / $70","$500",744.88,1601.49,1452.51,2309.12),
("ADP","OA EPO HDHP 6000-80%","Aetna OA EPO - no out-of-network","Y",6000,12000,6500,13000,0.80,"20% ded","20% ded","$10d / $45d / $70d","ded + 20%",437.90,941.49,853.91,1357.49),
("ADP","OA EPO 3000-80%","Aetna OA EPO - no out-of-network","N",3000,6000,7000,14000,0.80,"$30","$60","$10 / $50 / $80","$500",629.01,1352.36,1226.56,1949.92),
("ADP","OA EPO 8000-80% ValueINT","Aetna OA EPO - no out-of-network","N",8000,16000,10000,20000,0.80,"$10","$110","$10 / $50d / $90d","ded + 20%",374.03,804.17,729.36,1159.51),
# ---------------- TriNet / Aetna  (16 plans, strategy appendix pp.13-15) ----------------
("TriNet","Aetna ACO 6500 AZ","Banner Health Network","N",6500,13000,7500,15000,1.00,"$25","$65","$10 / $45 / $80","0% ded",385,889,785,1166),
("TriNet","Aetna PPO 7150","Aetna Managed Choice POS (Open Access)","N",7150,14300,7600,15200,1.00,"$40","$40 ded","$15 / $55 / $90","0% ded",400,924,816,1212),
("TriNet","Aetna HDHP 6350","Aetna Managed Choice POS (Open Access)","Y",6350,12700,6350,12700,1.00,"0% ded","0% ded","0% ded (all tiers)","0% ded",410,947,836,1242),
("TriNet","Aetna HDHP 4000","Aetna Managed Choice POS (Open Access)","Y",4000,8000,6850,13700,0.80,"20% ded","20% ded","$10d / $45d / $70d","20% ded",494,1140,1007,1495),
("TriNet","Aetna ACO 2500 AZ","Banner Health Network","N",2500,5000,7500,15000,0.80,"$30","$60","$10 / $45 / $80","$500",520,1202,1061,1576),
("TriNet","Aetna PPO 5000","Aetna Managed Choice POS (Open Access)","N",5000,10000,7600,15200,0.70,"$40","$80","$15 / $55 / $90","$350 ded",530,1223,1080,1603),
("TriNet","Aetna EPO 2000","Aetna Elect Choice EPO - no out-of-network","N",2000,4000,6500,13000,0.70,"$35","$70","$10 / $45 / $70","$500",581,1342,1185,1760),
("TriNet","Aetna ACO 1000 AZ","Banner Health Network","N",1000,2000,5500,11000,0.80,"$25","$50","$10 / $45 / $80","$500",599,1382,1220,1812),
("TriNet","Aetna PPO 3000","Aetna Managed Choice POS (Open Access)","N",3000,6000,6000,12000,1.00,"$30","$60","$10 / $50 / $80","$500",626,1446,1277,1897),
("TriNet","Aetna PPO 2000","Aetna Managed Choice POS (Open Access)","N",2000,4000,6850,13700,0.80,"$30","$60","$10 / $45 / $70","$500",632,1460,1290,1915),
("TriNet","Aetna EPO 1000","Aetna Elect Choice EPO - no out-of-network","N",1000,2000,5000,10000,0.70,"$30","$60","$10 / $45 / $70","$500",640,1477,1305,1938),
("TriNet","Aetna PPO 1000","Aetna Managed Choice POS (Open Access)","N",1000,2000,4500,9000,0.80,"$25","$50","$10 / $45 / $70","$500",670,1548,1367,2030),
("TriNet","Aetna HDHP 2000","Aetna Managed Choice POS (Open Access)","Y",2000,4000,3500,7000,1.00,"$30 ded","$60 ded","$10d / $45d / $70d","$350 ded",590,1363,1203,1787),
("TriNet","Aetna PPO 750","Aetna Managed Choice POS (Open Access)","N",750,1500,4000,8000,0.90,"$25","$50","$10 / $45 / $70","$350",743,1717,1516,2252),
("TriNet","Aetna EPO 0","Aetna Elect Choice EPO - no out-of-network","N",0,0,3000,6000,1.00,"$20","$40","$10 / $45 / $70","$350",794,1833,1619,2404),
("TriNet","Aetna PPO 300","Aetna Managed Choice POS (Open Access)","N",300,600,3000,6000,0.90,"$20","$40","$10 / $45 / $70","$350",863,1993,1760,2614),
# ---------------- Angle Health / CapFi  (3 plans, level funded) ----------------
("Angle","ANG HDHP 3400/5000","Cigna PPO (OON 50% after separate ded)","Y",3400,6800,5000,10000,0.80,"20% ded","20% ded","20% ded (all tiers)","20% ded",439.03,921.97,834.16,1361.00),
("Angle","ANG TRAD 2000/4000","Cigna PPO (OON 50% after separate ded)","N",2000,4000,4000,8000,0.80,"$20","$50","$20 / $60","$250",492.11,1033.43,934.01,1525.54),
("Angle","ANG TRAD 1000/2000","Cigna PPO (OON 50% after separate ded)","N",1000,2000,2000,4000,0.80,"$10","$30","$10 / $30","$200",546.33,1147.29,1038.02,1693.61),
]
FLOORS = [("ADP",0.50,374.03,"OA EPO 8000-80% ValueINT"),
          ("TriNet",0.70,385.00,"Aetna ACO 6500 AZ"),
          ("Angle",0.50,439.03,"ANG HDHP 3400/5000")]
TIERS = [("Employee only","EE",300.0),("Employee + spouse","ES",600.0),
         ("Employee + child(ren)","EC",600.0),("Family","FAM",1000.0)]
