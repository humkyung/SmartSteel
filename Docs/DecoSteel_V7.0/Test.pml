Default delete_log = "yes";
origin prompt = "Define end Point";
assign endx=%%point_x, var_type="float";
assign endy=%%point_y, var_type="float";
assign endz=%%point_z, var_type="float";
origin local = endx, endy, endz;
assign cd = 0.25, var_type = "float";
assign cw = 0.25, var_type = "float";
assign cwt = 0.009, var_type = "float";
assign cft = 0.014, var_type = "float";
assign BD1 = 0.244, var_type = "float";
assign BW1 = 0.175, var_type = "float";
assign Bwt1 = 0.007, var_type = "float";
assign Bft1 = 0.011, var_type = "float";
assign EPL1 = 0.48, var_type = "float";
assign EPW1 = 0.175, var_type = "float";
assign EPT1 = 0.019, var_type = "float";
assign EPL12 = 0.216, var_type = "float";
assign EPL13 = 0.02, var_type = "float";
assign STt1 =0.011, var_type = "float";
assign BD2 = 0.244, var_type = "float";
assign BW2 = 0.175, var_type = "float";
assign Bwt2 = 0.007, var_type = "float";
assign Bft2 = 0.011, var_type = "float";
assign EPL2 = 0.48, var_type = "float";
assign EPW2 = 0.175, var_type = "float";
assign EPT2 = 0.019, var_type = "float";
assign EPL22 = 0.216, var_type = "float";
assign EPL23 = 0.02, var_type = "float";
assign STt2 = 0.011, var_type = "float";
assign SPA_BOT1 = EPL1-EPL13, var_type = "float";
assign Px1 = cd/2+EPT1, var_type = "float";
assign Py1 = EPW1/2, var_type = "float";
assign Pzt1 = 0, var_type = "float";
assign Pzb1 = -EPL1, var_type = "float";
assign SPl = cd-(cft*2), var_type = "float";
assign SPw = (cw-cwt)/2, var_type = "float";
assign SPX11 = spl/2, var_type = "float";
assign SPY11 = cwt/2, var_type = "float";
assign SPY12 = (cwt/2)+spw, var_type = "float";
assign SPZ11 = SPA_BOT1+STt1, var_type = "float";
assign SPZ12 = SPA_BOT1, var_type = "float";
assign BRAX11 = PX1, var_type = "float";
assign BRAX12 = PX1, var_type = "float";
assign BRAX13 = PX1+1.75*EPL12, var_type = "float";
assign BRAX14 = PX1+1.75*EPL12, var_type = "float";
assign BRAY11 = -(BW1/2), var_type = "float";
assign BRAY12 = BW1/2, var_type = "float";
assign BRAY13 = BW1/2, var_type = "float";
assign BRAY14 = -(BW1/2), var_type = "float";
assign BRAZ11 = -SPZ12, var_type = "float";
assign BRAZ12 = -SPZ12, var_type = "float";
assign BRAZ13 = -SPZ12+EPL12, var_type = "float";
assign BRAZ14 = -SPZ12+EPL12, var_type = "float";
assign BRAX15 = PX1, var_type = "float";
assign BRAX16 = PX1, var_type = "float";
assign BRAX17 = PX1+1.75*EPL12, var_type = "float";
assign BRAY15 = -STt1/2, var_type = "float";
assign BRAY16 = -STt1/2, var_type = "float";
assign BRAY17 = -STt1/2, var_type = "float";
assign BRAZ15 = -SPZ12, var_type = "float";
assign BRAZ16 = -SPZ12+EPL12, var_type = "float";
assign BRAZ17 = -SPZ12+EPL12, var_type = "float";
assign BRA_STIFF_X1 = PX1+1.75*EPL12, var_type = "float";
assign BRA_STIFF_Y11 = -BW1/2, var_type = "float";
assign BRA_STIFF_Y12 = -Bwt1/2, var_type = "float";
assign BRA_STIFF_Y13 = -Bwt1/2, var_type = "float";
assign BRA_STIFF_Y14 = -BW1/2, var_type = "float";
assign BRA_STIFF_Y15 = BW1/2, var_type = "float";
assign BRA_STIFF_Y16 = Bwt1/2, var_type = "float";
assign BRA_STIFF_Y17 = Bwt1/2, var_type = "float";
assign BRA_STIFF_Y18 = BW1/2, var_type = "float";
assign BRA_STIFF_Z11 = -BD1+Bft1, var_type = "float";
assign BRA_STIFF_Z12 = -BD1+Bft1, var_type = "float";
assign BRA_STIFF_Z13 = BRA_STIFF_Z11+(BD1/2)-BFT1, var_type = "float";
assign BRA_STIFF_Z14 = BRA_STIFF_Z12+(BD1/2)-BFT1, var_type = "float";
assign SPA_BOT2 = EPL2-EPL23, var_type = "float";
assign Px2 = cd/2, var_type = "float";
assign Py2 = EPW2/2, var_type = "float";
assign Pzt2 = 0, var_type = "float";
assign Pzb2 = -EPL2, var_type = "float";
assign SPZ21 = SPA_BOT2+STt2, var_type = "float";
assign SPZ22 = SPA_BOT2, var_type = "float";
assign BRAX21 = PX2+EPT2, var_type = "float";
assign BRAX22 = PX2+EPT2, var_type = "float";
assign BRAX23 = PX2+1.75*EPL22-EPT2, var_type = "float";
assign BRAX24 = PX2+1.75*EPL22-EPT2, var_type = "float";
assign BRAY21 = BW2/2, var_type = "float";
assign BRAY22 = -(BW2/2), var_type = "float";
assign BRAY23 = -(BW2/2), var_type = "float";
assign BRAY24 = BW2/2, var_type = "float";
assign BRAZ21 = -SPZ22, var_type = "float";
assign BRAZ22 = -SPZ22, var_type = "float";
assign BRAZ23 = -SPZ22+EPL22, var_type = "float";
assign BRAZ24 = -SPZ22+EPL22, var_type = "float";
assign BRAX25 = PX2+EPT2, var_type = "float";
assign BRAX26 = PX2+EPT2, var_type = "float";
assign BRAX27 = PX2+1.75*EPL22-EPT2, var_type = "float";
assign BRAY25 = STt2/2, var_type = "float";
assign BRAY26 = STt2/2, var_type = "float";
assign BRAY27 = STt2/2, var_type = "float";
assign BRAZ25 = -SPZ22, var_type = "float";
assign BRAZ26 = -SPZ22+EPL22, var_type = "float";
assign BRAZ27 = -SPZ22+EPL22, var_type = "float";
assign BRA_STIFF_X2 = PX2+1.75*EPL22-EPT2, var_type = "float";
assign BRA_STIFF_Y21 = -BW2/2, var_type = "float";
assign BRA_STIFF_Y22 = -Bwt2/2, var_type = "float";
assign BRA_STIFF_Y23 = -Bwt2/2, var_type = "float";
assign BRA_STIFF_Y24 = -BW2/2, var_type = "float";
assign BRA_STIFF_Y25 = BW2/2, var_type = "float";
assign BRA_STIFF_Y26 = Bwt2/2, var_type = "float";
assign BRA_STIFF_Y27 = Bwt2/2, var_type = "float";
assign BRA_STIFF_Y28 = BW2/2, var_type = "float";
assign BRA_STIFF_Z21 = -BD2+Bft2, var_type = "float";
assign BRA_STIFF_Z22 = -BD2+Bft2, var_type = "float";
assign BRA_STIFF_Z23 = BRA_STIFF_Z21+(BD2/2)-BFT2, var_type = "float";
assign BRA_STIFF_Z24 = BRA_STIFF_Z22+(BD2/2)-BFT2, var_type = "float";
plc_area
vert1 = Py1, Px1, Pzt1,
vert2 = Py1, Px1, Pzb1,
vert3 = -Py1, Px1, Pzb1,
vert4 = -Py1, Px1, Pzt1,
class = 3, grade = "A36", material = "Steel", name = "EP_0.019", 
thickness = EPT1;
plc_area
vert1 = Py2, -Px2, Pzt2,
vert2 = Py2, -Px2, Pzb2,
vert3 = -Py2, -Px2, Pzb2,
vert4 = -Py2, -Px2, Pzt2,
class = 3, grade = "A36", material = "Steel", name = "EP_0.019", 
thickness = EPT2;
plc_area
vert1 = spy11, spx11, -STt1,
vert2 = spy12, spx11, -STt1,
vert3 = spy12, -spx11, -STt1,
vert4 = spy11, -spx11, -STt1,
class = 3, grade = "A36", material = "Steel", name = "SP_0.011", 
thickness = STt1;
plc_area
vert1 = -spy11, spx11, 0,
vert2 = -spy12, spx11, 0,
vert3 = -spy12, -spx11, 0,
vert4 = -spy11, -spx11, 0,
class = 3, grade = "A36", material = "Steel", name = "SP_0.011", 
thickness = STt1;
plc_area
vert1 = spy11, spx11, -SPZ11,
vert2 = spy12, spx11, -SPZ11,
vert3 = spy12, -spx11, -SPZ11,
vert4 = spy11, -spx11, -SPZ11,
class = 3, grade = "A36", material = "Steel", name = "SP_0.011", 
thickness = STt1;
plc_area
vert1 = -spy11, spx11, -SPZ12,
vert2 = -spy12, spx11, -SPZ12,
vert3 = -spy12, -spx11, -SPZ12,
vert4 = -spy11, -spx11, -SPZ12,
class = 3, grade = "A36", material = "Steel", name = "SP_0.011", 
thickness = STt1;
plc_area
vert1 = BRAy11, BRAx11, BRAZ11,
vert2 = BRAy12, BRAx12, BRAZ12,
vert3 = BRAy13, BRAx13, BRAZ13,
vert4 = BRAy14, BRAx14, BRAZ14,
class = 3, grade = "A36", material = "Steel", name = "SP_0.011", 
thickness = STt1;
plc_area
vert1 = BRAy15, BRAx15, BRAZ14,
vert2 = BRAy16, BRAx16, BRAZ15,
vert3 = BRAy17, BRAx17, BRAZ16,
class = 3, grade = "A36", material = "Steel", name = "SP_0.011", 
thickness = STt1;
plc_area
vert1 = BRA_STIFF_Y11, BRA_STIFF_X1, BRA_STIFF_Z11,
vert2 = BRA_STIFF_Y12, BRA_STIFF_X1, BRA_STIFF_Z12,
vert3 = BRA_STIFF_Y13, BRA_STIFF_X1,  BRA_STIFF_Z13,
vert4 = BRA_STIFF_Y14, BRA_STIFF_X1, BRA_STIFF_Z14,
class = 3, grade = "A36", material = "Steel", name = "SP_0.011", 
thickness = STt1;
plc_area
vert1 = BRA_STIFF_Y15, BRA_STIFF_X1+STT1, BRA_STIFF_Z11,
vert2 = BRA_STIFF_Y16, BRA_STIFF_X1+STT1, BRA_STIFF_Z12,
vert3 = BRA_STIFF_Y17, BRA_STIFF_X1+STT1, BRA_STIFF_Z13,
vert4 = BRA_STIFF_Y18, BRA_STIFF_X1+STT1, BRA_STIFF_Z14,
class = 3, grade = "A36", material = "Steel", name = "SP_0.011", 
thickness = STt1;
plc_area
vert1 = BRAy21, -BRAx21, BRAZ21,
vert2 = BRAy22, -BRAx22, BRAZ22,
vert3 = BRAy23, -BRAx23, BRAZ23,
vert4 = BRAy24, -BRAx24, BRAZ24,
class = 3, grade = "A36", material = "Steel", name = "SP_0.011", 
thickness = STt2;
plc_area
vert1 = BRAy25, -BRAx25, BRAZ24,
vert2 = BRAy26, -BRAx26, BRAZ25,
vert3 = BRAy27, -BRAx27, BRAZ26,
class = 3, grade = "A36", material = "Steel", name = "SP_0.011", 
thickness = STt2;
plc_area
vert1 = BRA_STIFF_Y21, -(BRA_STIFF_X2+STT2), BRA_STIFF_Z21,
vert2 = BRA_STIFF_Y22, -(BRA_STIFF_X2+STT2), BRA_STIFF_Z22,
vert3 = BRA_STIFF_Y23, -(BRA_STIFF_X2+STT2), BRA_STIFF_Z23,
vert4 = BRA_STIFF_Y24, -(BRA_STIFF_X2+STT2), BRA_STIFF_Z24,
class = 3, grade = "A36", material = "Steel", name = "SP_0.011", 
thickness = STt2;
plc_area
vert1 = BRA_STIFF_Y25, -BRA_STIFF_X2, BRA_STIFF_Z21,
vert2 = BRA_STIFF_Y26, -BRA_STIFF_X2, BRA_STIFF_Z22,
vert3 = BRA_STIFF_Y27, -BRA_STIFF_X2, BRA_STIFF_Z23,
vert4 = BRA_STIFF_Y28, -BRA_STIFF_X2, BRA_STIFF_Z24,
class = 3, grade = "A36", material = "Steel", name = "SP_0.011", 
thickness = STt2;
plc_area
vert1 = -0.038, 0.160,-0.066,
vert2 = -0.038, 0.160,-0.084,
vert3 = -0.052, 0.160,-0.092,
vert4 = -0.067, 0.160,-0.084,
vert5 = -0.068, 0.160,-0.066,
vert6 = -0.053, 0.160,-0.058,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.067, 0.160,-0.066,
vert2 = 0.067, 0.160,-0.084,
vert3 = 0.053, 0.160,-0.092,
vert4 = 0.038, 0.160,-0.084,
vert5 = 0.037, 0.160,-0.066,
vert6 = 0.052, 0.160,-0.058,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = -0.038, 0.160,-0.126,
vert2 = -0.038, 0.160,-0.144,
vert3 = -0.052, 0.160,-0.152,
vert4 = -0.067, 0.160,-0.144,
vert5 = -0.068, 0.160,-0.126,
vert6 = -0.053, 0.160,-0.118,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.067, 0.160,-0.126,
vert2 = 0.067, 0.160,-0.144,
vert3 = 0.053, 0.160,-0.152,
vert4 = 0.038, 0.160,-0.144,
vert5 = 0.037, 0.160,-0.126,
vert6 = 0.052, 0.160,-0.118,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = -0.038, 0.160,-0.186,
vert2 = -0.038, 0.160,-0.204,
vert3 = -0.052, 0.160,-0.212,
vert4 = -0.067, 0.160,-0.204,
vert5 = -0.068, 0.160,-0.186,
vert6 = -0.053, 0.160,-0.178,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.067, 0.160,-0.186,
vert2 = 0.067, 0.160,-0.204,
vert3 = 0.053, 0.160,-0.212,
vert4 = 0.038, 0.160,-0.204,
vert5 = 0.037, 0.160,-0.186,
vert6 = 0.052, 0.160,-0.178,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = -0.038, 0.160,-0.291,
vert2 = -0.038, 0.160,-0.309,
vert3 = -0.052, 0.160,-0.317,
vert4 = -0.067, 0.160,-0.309,
vert5 = -0.068, 0.160,-0.291,
vert6 = -0.053, 0.160,-0.283,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.067, 0.160,-0.291,
vert2 = 0.067, 0.160,-0.309,
vert3 = 0.053, 0.160,-0.317,
vert4 = 0.038, 0.160,-0.309,
vert5 = 0.037, 0.160,-0.291,
vert6 = 0.052, 0.160,-0.283,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = -0.038, 0.160,-0.366,
vert2 = -0.038, 0.160,-0.384,
vert3 = -0.052, 0.160,-0.392,
vert4 = -0.067, 0.160,-0.384,
vert5 = -0.068, 0.160,-0.366,
vert6 = -0.053, 0.160,-0.358,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.067, 0.160,-0.366,
vert2 = 0.067, 0.160,-0.384,
vert3 = 0.053, 0.160,-0.392,
vert4 = 0.038, 0.160,-0.384,
vert5 = 0.037, 0.160,-0.366,
vert6 = 0.052, 0.160,-0.358,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = -0.038, -0.144,-0.066,
vert2 = -0.038, -0.144,-0.084,
vert3 = -0.052, -0.144,-0.092,
vert4 = -0.067, -0.144,-0.084,
vert5 = -0.068, -0.144,-0.066,
vert6 = -0.053, -0.144,-0.058,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.067, -0.144,-0.066,
vert2 = 0.067, -0.144,-0.084,
vert3 = 0.053, -0.144,-0.092,
vert4 = 0.038, -0.144,-0.084,
vert5 = 0.037, -0.144,-0.066,
vert6 = 0.052, -0.144,-0.058,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = -0.038, -0.144,-0.126,
vert2 = -0.038, -0.144,-0.144,
vert3 = -0.052, -0.144,-0.152,
vert4 = -0.067, -0.144,-0.144,
vert5 = -0.068, -0.144,-0.126,
vert6 = -0.053, -0.144,-0.118,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.067, -0.144,-0.126,
vert2 = 0.067, -0.144,-0.144,
vert3 = 0.053, -0.144,-0.152,
vert4 = 0.038, -0.144,-0.144,
vert5 = 0.037, -0.144,-0.126,
vert6 = 0.052, -0.144,-0.118,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = -0.038, -0.144,-0.186,
vert2 = -0.038, -0.144,-0.204,
vert3 = -0.052, -0.144,-0.212,
vert4 = -0.067, -0.144,-0.204,
vert5 = -0.068, -0.144,-0.186,
vert6 = -0.053, -0.144,-0.178,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.067, -0.144,-0.186,
vert2 = 0.067, -0.144,-0.204,
vert3 = 0.053, -0.144,-0.212,
vert4 = 0.038, -0.144,-0.204,
vert5 = 0.037, -0.144,-0.186,
vert6 = 0.052, -0.144,-0.178,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = -0.038, -0.144,-0.291,
vert2 = -0.038, -0.144,-0.309,
vert3 = -0.052, -0.144,-0.317,
vert4 = -0.067, -0.144,-0.309,
vert5 = -0.068, -0.144,-0.291,
vert6 = -0.053, -0.144,-0.283,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.067, -0.144,-0.291,
vert2 = 0.067, -0.144,-0.309,
vert3 = 0.053, -0.144,-0.317,
vert4 = 0.038, -0.144,-0.309,
vert5 = 0.037, -0.144,-0.291,
vert6 = 0.052, -0.144,-0.283,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = -0.038, -0.144,-0.366,
vert2 = -0.038, -0.144,-0.384,
vert3 = -0.052, -0.144,-0.392,
vert4 = -0.067, -0.144,-0.384,
vert5 = -0.068, -0.144,-0.366,
vert6 = -0.053, -0.144,-0.358,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.067, -0.144,-0.366,
vert2 = 0.067, -0.144,-0.384,
vert3 = 0.053, -0.144,-0.392,
vert4 = 0.038, -0.144,-0.384,
vert5 = 0.037, -0.144,-0.366,
vert6 = 0.052, -0.144,-0.358,
class = 3, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
