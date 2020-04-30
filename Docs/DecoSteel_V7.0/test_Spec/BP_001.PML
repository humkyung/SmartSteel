Default delete_log = "yes";
origin prompt = "Define end Point";
assign endx=%%point_x, var_type="Float";
assign endy=%%point_y, var_type="Float";
assign endz=%%point_z, var_type="Float";
origin local = endx, endy, endz;
assign HBd = 0.2, var_type="Float";
assign HBw = 0.2, var_type="Float";
assign HBwt = 0.008, var_type="Float";
assign HBft = 0.012, var_type="Float";
assign Lx = 0.4, var_type="Float";
assign Ly = 0.3, var_type="Float";
assign Lz = 0.016, var_type="Float";
assign RBt = 0.008, var_type="Float";
assign RBr = 0.01, var_type="Float";
assign RBe = 0.01, var_type="Float";
assign DimF = 0.05, var_type="Float";
assign RBsw = 0.1, var_type="Float";
assign RBww = 0.05, var_type="Float";
assign RBh = 0.135, var_type="Float";
assign Lhx = 0.2, var_type="Float";
assign Lhy = 0.15, var_type="Float";
assign RBtH = 0.004, var_type="Float";
assign WRY11 =  HBw/2 , var_type="Float";
assign WRZ11 =  Lz, var_type="Float";
assign WRY12 =  WRY11+RBww-RBe , var_type="Float";
assign WRZ12 =  Lz, var_type="Float";
assign WRY13 =  WRY11+RBww-RBe , var_type="Float";
assign WRZ13 =  Lz+RBr , var_type="Float";
assign WRY14 =  WRY11+RBr , var_type="Float";
assign WRZ14 =  Lz+RBh , var_type="Float";
assign WRY15 =  HBw/2, var_type="Float";
assign WRZ15 =  Lz+RBh , var_type="Float";
assign WRX1 =  HBd/2 , var_type="Float";
assign WRY21 =  -HBw/2 , var_type="Float";
assign WRZ21 =  Lz, var_type="Float";
assign WRY22 =  -WRY11-RBww+RBe , var_type="Float";
assign WRZ22 =  Lz, var_type="Float";
assign WRY23 =  -WRY11-RBww+RBe , var_type="Float";
assign WRZ23 =  Lz+RBr , var_type="Float";
assign WRY24 =  -WRY11-RBr , var_type="Float";
assign WRZ24 =  Lz+RBh , var_type="Float";
assign WRY25 =  -HBw/2, var_type="Float";
assign WRZ25 =  Lz+RBh , var_type="Float";
assign WRX2 =  HBd/2 , var_type="Float";
assign SRX11 =  HBd/2 , var_type="Float";
assign SRZ11 =  Lz, var_type="Float";
assign SRX12 =  SRX11+RBsw-RBe , var_type="Float";
assign SRZ12 =  Lz, var_type="Float";
assign SRX13 =  SRX11+RBsw-RBe , var_type="Float";
assign SRZ13 =  Lz+RBr , var_type="Float";
assign SRX14 =  SRX11+RBr , var_type="Float";
assign SRZ14 =  Lz+RBh , var_type="Float";
assign SRX15 =  HBd/2, var_type="Float";
assign SRZ15 =  Lz+RBh , var_type="Float";
assign SRY1 =  DimF , var_type="Float";
assign SRX21 =  -HBd/2 , var_type="Float";
assign SRZ21 =  Lz, var_type="Float";
assign SRX22 =  -SRX11-RBSw+RBe , var_type="Float";
assign SRZ22 =  Lz, var_type="Float";
assign SRX23 =  -SRX11-RBsw+RBe , var_type="Float";
assign SRZ23 =  Lz+RBr , var_type="Float";
assign SRX24 =  -SRX11-RBr , var_type="Float";
assign SRZ24 =  Lz+RBh , var_type="Float";
assign SRX25 =  -HBd/2, var_type="Float";
assign SRZ25 =  Lz+RBh , var_type="Float";
assign SRY2 =  DimF , var_type="Float";
plc_area
vert1 = Lhx, Lhy, 0,
vert2 = -Lhx, Lhy, 0,
vert3 = -Lhx, -Lhy, 0,
vert4 = Lhx,-Lhy, 0,
class = 3, grade = "A36", material = "Steel", name = "BP_0.016", 
thickness = Lz;
plc_area
vert1 = WRX1-RBt, WRY11, WRZ11,
vert2 = WRX1-RBt, WRY12, WRZ12,
vert3 = WRX1-RBt, WRY13, WRZ13,
vert4 = WRX1-RBt, WRY14, WRZ14,
vert5 = WRX1-RBt, WRY15, WRZ15,
class = 3, grade = "A36", material = "Steel", name = "RP_0.008", 
thickness = RBt;
plc_area
vert1 = WRX2, WRY21, WRZ21,
vert2 = WRX2, WRY22, WRZ22,
vert3 = WRX2, WRY23, WRZ23,
vert4 = WRX2, WRY24, WRZ24,
vert5 = WRX2, WRY25, WRZ25,
class = 3, grade = "A36", material = "Steel", name = "RP_0.008", 
thickness = RBt;
plc_area
vert1 = -WRX2+RBt, WRY21, WRZ21,
vert2 = -WRX2+RBt, WRY22, WRZ22,
vert3 = -WRX2+RBt, WRY23, WRZ23,
vert4 = -WRX2+RBt, WRY24, WRZ24,
vert5 = -WRX2+RBt, WRY25, WRZ25,
class = 3, grade = "A36", material = "Steel", name = "RP_0.008", 
thickness = RBt;
plc_area
vert1 = -WRX1, WRY11, WRZ11,
vert2 = -WRX1, WRY12, WRZ12,
vert3 = -WRX1, WRY13, WRZ13,
vert4 = -WRX1, WRY14, WRZ14,
vert5 = -WRX1, WRY15, WRZ15,
class = 3, grade = "A36", material = "Steel", name = "RP_0.008", 
thickness = RBt;
plc_area
vert1 = SRX11, SRY1+RBtH, SRZ11,
vert2 = SRX12, SRY1+RBtH, SRZ12,
vert3 = SRX13, SRY1+RBtH, SRZ13,
vert4 = SRX14, SRY1+RBtH, SRZ14,
vert5 = SRX15, SRY1+RBtH, SRZ15,
class = 3, grade = "A36", material = "Steel", name = "RP_0.008", 
thickness = RBt;
plc_area
vert1 = SRX21, SRY2-RBtH, SRZ21,
vert2 = SRX22, SRY2-RBtH, SRZ22,
vert3 = SRX23, SRY2-RBtH, SRZ23,
vert4 = SRX24, SRY2-RBtH, SRZ24,
vert5 = SRX25, SRY2-RBtH, SRZ25,
class = 3, grade = "A36", material = "Steel", name = "RP_0.008", 
thickness = RBt;
plc_area
vert1 = SRX11, -SRY1+RBtH, SRZ11,
vert2 = SRX12, -SRY1+RBtH, SRZ12,
vert3 = SRX13, -SRY1+RBtH, SRZ13,
vert4 = SRX14, -SRY1+RBtH, SRZ14,
vert5 = SRX15, -SRY1+RBtH, SRZ15,
class = 3, grade = "A36", material = "Steel", name = "RP_0.008", 
thickness = RBt;
plc_area
vert1 = SRX21, -SRY2-RBtH, SRZ21,
vert2 = SRX22, -SRY2-RBtH, SRZ22,
vert3 = SRX23, -SRY2-RBtH, SRZ23,
vert4 = SRX24, -SRY2-RBtH, SRZ24,
vert5 = SRX25, -SRY2-RBtH, SRZ25,
class = 3, grade = "A36", material = "Steel", name = "RP_0.008", 
thickness = RBt;
plc_area
vert1 = 0.162, -0.093,0.029,
vert2 = 0.162, -0.107,0.029,
vert3 = 0.150, -0.114,0.029,
vert4 = 0.138, -0.107,0.029,
vert5 = 0.138, -0.093,0.029,
vert6 = 0.150, -0.086,0.029,
class = 3, grade = "A36", material = "Steel", name = "AB_M16", 
thickness = 0.013;
plc_area
vert1 = 0.162, 0.007,0.029,
vert2 = 0.162, -0.007,0.029,
vert3 = 0.150, -0.014,0.029,
vert4 = 0.138, -0.007,0.029,
vert5 = 0.138, 0.007,0.029,
vert6 = 0.150, 0.014,0.029,
class = 3, grade = "A36", material = "Steel", name = "AB_M16", 
thickness = 0.013;
plc_area
vert1 = 0.162, 0.107,0.029,
vert2 = 0.162, 0.093,0.029,
vert3 = 0.150, 0.086,0.029,
vert4 = 0.138, 0.093,0.029,
vert5 = 0.138, 0.107,0.029,
vert6 = 0.150, 0.114,0.029,
class = 3, grade = "A36", material = "Steel", name = "AB_M16", 
thickness = 0.013;
plc_area
vert1 = -0.138, -0.093,0.029,
vert2 = -0.138, -0.107,0.029,
vert3 = -0.150, -0.114,0.029,
vert4 = -0.162, -0.107,0.029,
vert5 = -0.162, -0.093,0.029,
vert6 = -0.150, -0.086,0.029,
class = 3, grade = "A36", material = "Steel", name = "AB_M16", 
thickness = 0.013;
plc_area
vert1 = -0.138, 0.007,0.029,
vert2 = -0.138, -0.007,0.029,
vert3 = -0.150, -0.014,0.029,
vert4 = -0.162, -0.007,0.029,
vert5 = -0.162, 0.007,0.029,
vert6 = -0.150, 0.014,0.029,
class = 3, grade = "A36", material = "Steel", name = "AB_M16", 
thickness = 0.013;
plc_area
vert1 = -0.138, 0.107,0.029,
vert2 = -0.138, 0.093,0.029,
vert3 = -0.150, 0.086,0.029,
vert4 = -0.162, 0.093,0.029,
vert5 = -0.162, 0.107,0.029,
vert6 = -0.150, 0.114,0.029,
class = 3, grade = "A36", material = "Steel", name = "AB_M16", 
thickness = 0.013;
