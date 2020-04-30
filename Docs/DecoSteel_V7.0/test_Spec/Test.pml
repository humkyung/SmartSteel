Default delete_log = "yes";
origin prompt = "Define end Point";
assign endx=%%point_x, var_type="float";
assign endy=%%point_y, var_type="float";
assign endz=%%point_z, var_type="float";
origin local = endx, endy, endz;
assign cd = 0.2, var_type = "float";
assign cw = 0.2, var_type = "float";
assign cwt = 0.008, var_type = "float";
assign cft = 0.012, var_type = "float";
assign BD = 0.148, var_type = "float";
assign BW = 0.1, var_type = "float";
assign Bwt = 0.006, var_type = "float";
assign Bft = 0.009, var_type = "float";
assign EPL = 0.3, var_type = "float";
assign EPW = 0.125, var_type = "float";
assign EPT = 0.019, var_type = "float";
assign EPL2 = 0.132, var_type = "float";
assign EPL3 = 0.02, var_type = "float";
assign STt = 0.009, var_type = "float";
assign SPA_BOT = EPL-EPL3, var_type = "float";
assign Px = cd/2, var_type = "float";
assign Py = EPW/2, var_type = "float";
assign Pzt = 0, var_type = "float";
assign Pzb = -EPL, var_type = "float";
assign SPl = cd-(cft*2), var_type = "float";
assign SPw = (cw-cwt)/2, var_type = "float";
assign SPX1 = spl/2, var_type = "float";
assign SPY1 = cwt/2, var_type = "float";
assign SPY2 = (cwt/2)+spw, var_type = "float";
assign SPZ1 = SPA_BOT-STt, var_type = "float";
assign SPZ2 = SPA_BOT, var_type = "float";
assign BRAX1 = PX+EPT, var_type = "float";
assign BRAX2 = PX+EPT, var_type = "float";
assign BRAX3 = PX+1.75*EPL2, var_type = "float";
assign BRAX4 = PX+1.75*EPL2, var_type = "float";
assign BRAY1 = -(BW/2), var_type = "float";
assign BRAY2 = BW/2, var_type = "float";
assign BRAY3 = BW/2, var_type = "float";
assign BRAY4 = -(BW/2), var_type = "float";
assign BRAZ1 = -SPZ2, var_type = "float";
assign BRAZ2 = -SPZ2, var_type = "float";
assign BRAZ3 = -SPZ2+EPL2, var_type = "float";
assign BRAZ4 = -SPZ2+EPL2, var_type = "float";
assign BRAX5 = PX+EPT, var_type = "float";
assign BRAX6 = PX+EPT, var_type = "float";
assign BRAX7 = PX+1.75*EPL2, var_type = "float";
assign BRAY5 = -STt/2, var_type = "float";
assign BRAY6 = -STt/2, var_type = "float";
assign BRAY7 = -STt/2, var_type = "float";
assign BRAZ5 = -SPZ2, var_type = "float";
assign BRAZ6 = -SPZ2+EPL2, var_type = "float";
assign BRAZ7 = -SPZ2+EPL2, var_type = "float";
assign BRA_STIFF_X = PX+1.75*EPL2, var_type = "float";
assign BRA_STIFF_Y1 = -BW/2, var_type = "float";
assign BRA_STIFF_Y2 = -Bwt/2, var_type = "float";
assign BRA_STIFF_Y3 = -Bwt/2, var_type = "float";
assign BRA_STIFF_Y4 = -BW/2, var_type = "float";
assign BRA_STIFF_Y5 = BW/2, var_type = "float";
assign BRA_STIFF_Y6 = Bwt/2, var_type = "float";
assign BRA_STIFF_Y7 = Bwt/2, var_type = "float";
assign BRA_STIFF_Y8 = BW/2, var_type = "float";
assign BRA_STIFF_Z1 = -BD+Bft, var_type = "float";
assign BRA_STIFF_Z2 = -BD+Bft, var_type = "float";
assign BRA_STIFF_Z3 = BRA_STIFF_Z1+(BD/2)-BFT, var_type = "float";
assign BRA_STIFF_Z4 = BRA_STIFF_Z2+(BD/2)-BFT, var_type = "float";
plc_area
vert1 = -Px, Py, Pzt,
vert2 = -Px, Py, Pzb,
vert3 = -Px, -Py, Pzb,
vert4 = -Px, -Py, Pzt,
class = 1, grade = "A36", material = "Steel", name = "EP_0.019", 
thickness = EPT;
plc_area
vert1 = spx1, spy1, 0,
vert2 = -spx1, spy1, 0,
vert3 = -spx1, spy2, 0,
vert4 = spx1, spy2, 0,
class = 1, grade = "A36", material = "Steel", name = "SP_0.009", 
thickness = STt;
plc_area
vert1 = spx1, -spy1, -STt,
vert2 = -spx1, -spy1, -STt,
vert3 = -spx1, -spy2, -STt,
vert4 = spx1, -spy2, -STt,
class = 1, grade = "A36", material = "Steel", name = "SP_0.009", 
thickness = STt;
plc_area
vert1 = spx1, spy1, -SPZ1,
vert2 = -spx1, spy1, -SPZ1,
vert3 = -spx1, spy2, -SPZ1,
vert4 = spx1, spy2, -SPZ1,
class = 1, grade = "A36", material = "Steel", name = "SP_0.009", 
thickness = STt;
plc_area
vert1 = spx1, -spy1, -SPZ2,
vert2 = -spx1, -spy1, -SPZ2,
vert3 = -spx1, -spy2, -SPZ2,
vert4 = spx1, -spy2, -SPZ2,
class = 1, grade = "A36", material = "Steel", name = "SP_0.009", 
thickness = STt;
plc_area
vert1 = -BRAx1, BRAy1, BRAZ1,
vert2 = -BRAx2, BRAy2, BRAZ2,
vert3 = -BRAx3, BRAy3, BRAZ3,
vert4 = -BRAx4, BRAy4, BRAZ4,
class = 1, grade = "A36", material = "Steel", name = "SP_0.009", 
thickness = STt;
plc_area
vert1 = -BRAx5, BRAY5, BRAZ4,
vert2 = -BRAx6, BRAY6, BRAZ5,
vert3 = -BRAx7, BRAY7, BRAZ6,
class = 1, grade = "A36", material = "Steel", name = "SP_0.009", 
thickness = STt;
plc_area
vert1 = -BRA_STIFF_X, BRA_STIFF_Y1, BRA_STIFF_Z1,
vert2 = -BRA_STIFF_X, BRA_STIFF_Y2, BRA_STIFF_Z2,
vert3 = -BRA_STIFF_X, BRA_STIFF_Y3, BRA_STIFF_Z3,
vert4 = -BRA_STIFF_X, BRA_STIFF_Y4, BRA_STIFF_Z4,
class = 1, grade = "A36", material = "Steel", name = "SP_0.009", 
thickness = STt;
plc_area
vert1 = -(BRA_STIFF_X-STT), BRA_STIFF_Y5, BRA_STIFF_Z1,
vert2 = -(BRA_STIFF_X-STT), BRA_STIFF_Y6, BRA_STIFF_Z2,
vert3 = -(BRA_STIFF_X-STT), BRA_STIFF_Y7, BRA_STIFF_Z3,
vert4 = -(BRA_STIFF_X-STT), BRA_STIFF_Y8, BRA_STIFF_Z4,
class = 1, grade = "A36", material = "Steel", name = "SP_0.009", 
thickness = STt;
plc_area
vert1 = -0.119, -0.026,-0.068,
vert2 = -0.119, -0.025,-0.082,
vert3 = -0.119, -0.037,-0.089,
vert4 = -0.119, -0.049,-0.082,
vert5 = -0.119, -0.050,-0.068,
vert6 = -0.119, -0.038,-0.061,
class = 1, grade = "A36", material = "Steel", name = "HTB_M16", 
thickness = 0.013;
plc_area
vert1 = -0.119, 0.049,-0.068,
vert2 = -0.119, 0.050,-0.082,
vert3 = -0.119, 0.038,-0.089,
vert4 = -0.119, 0.026,-0.082,
vert5 = -0.119, 0.025,-0.068,
vert6 = -0.119, 0.037,-0.061,
class = 1, grade = "A36", material = "Steel", name = "HTB_M16", 
thickness = 0.013;
plc_area
vert1 = -0.119, -0.026,-0.208,
vert2 = -0.119, -0.025,-0.222,
vert3 = -0.119, -0.037,-0.229,
vert4 = -0.119, -0.049,-0.222,
vert5 = -0.119, -0.050,-0.208,
vert6 = -0.119, -0.038,-0.201,
class = 1, grade = "A36", material = "Steel", name = "HTB_M16", 
thickness = 0.013;
plc_area
vert1 = -0.119, 0.049,-0.208,
vert2 = -0.119, 0.050,-0.222,
vert3 = -0.119, 0.038,-0.229,
vert4 = -0.119, 0.026,-0.222,
vert5 = -0.119, 0.025,-0.208,
vert6 = -0.119, 0.037,-0.201,
class = 1, grade = "A36", material = "Steel", name = "HTB_M16", 
thickness = 0.013;
