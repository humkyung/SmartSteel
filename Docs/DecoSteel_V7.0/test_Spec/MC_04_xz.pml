Default delete_log = "yes";
origin prompt = "Define end Point";
assign endx=%%point_x, var_type="float";
assign endy=%%point_y, var_type="float";
assign endz=%%point_z, var_type="float";
origin local = endx, endy, endz;
assign cd = 0.198, var_type = "float";
assign cw = 0.099, var_type = "float";
assign cwt = 0.0045, var_type = "float";
assign cft = 0.007, var_type = "float";
assign BD1 = 0.6, var_type = "float";
assign BW1 = 0.2, var_type = "float";
assign Bwt1 = 0.011, var_type = "float";
assign Bft1 = 0.017, var_type = "float";
assign EPL1 = 0.94, var_type = "float";
assign EPW1 = 0.23, var_type = "float";
assign EPT1 = 0.028, var_type = "float";
assign EPL12 = 0.32, var_type = "float";
assign EPL13 = 0.02, var_type = "float";
assign STt1 = 0.012, var_type = "float";
assign BD2 = 0.194, var_type = "float";
assign BW2 = 0.15, var_type = "float";
assign Bwt2 = 0.006, var_type = "float";
assign Bft2 = 0.009, var_type = "float";
assign EPL2 = 0.4, var_type = "float";
assign EPW2 = 0.15, var_type = "float";
assign EPT2 = 0.019, var_type = "float";
assign EPL22 = 0.186, var_type = "float";
assign EPL23 = 0.02, var_type = "float";
assign STt2 = 0.009, var_type = "float";
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
assign SPZ11 = SPA_BOT1-STt1, var_type = "float";
assign SPZ12 = SPA_BOT1, var_type = "float";
assign BRAX11 = PX1, var_type = "float";
assign BRAX12 = PX1, var_type = "float";
assign BRAX13 = PX1+1.75*EPL12-EPT1, var_type = "float";
assign BRAX14 = PX1+1.75*EPL12-EPT1, var_type = "float";
assign BRAY11 = BW1/2, var_type = "float";
assign BRAY12 = -(BW1/2), var_type = "float";
assign BRAY13 = -(BW1/2), var_type = "float";
assign BRAY14 = BW1/2, var_type = "float";
assign BRAZ11 = -SPZ12, var_type = "float";
assign BRAZ12 = -SPZ12, var_type = "float";
assign BRAZ13 = -SPZ12+EPL12, var_type = "float";
assign BRAZ14 = -SPZ12+EPL12, var_type = "float";
assign BRAX15 = PX1, var_type = "float";
assign BRAX16 = PX1, var_type = "float";
assign BRAX17 = PX1+1.75*EPL12-EPT1, var_type = "float";
assign BRAY15 = STt1/2, var_type = "float";
assign BRAY16 = STt1/2, var_type = "float";
assign BRAY17 = STt1/2, var_type = "float";
assign BRAZ15 = -SPZ12, var_type = "float";
assign BRAZ16 = -SPZ12+EPL12, var_type = "float";
assign BRAZ17 = -SPZ12+EPL12, var_type = "float";
assign BRA_STIFF_X1 = PX1+1.75*EPL12-EPT1, var_type = "float";
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
assign SPZ21 = SPA_BOT2-STt2, var_type = "float";
assign SPZ22 = SPA_BOT2, var_type = "float";
assign BRAX21 = PX2+EPT2, var_type = "float";
assign BRAX22 = PX2+EPT2, var_type = "float";
assign BRAX23 = PX2+1.75*EPL22, var_type = "float";
assign BRAX24 = PX2+1.75*EPL22, var_type = "float";
assign BRAY21 = -(BW2/2), var_type = "float";
assign BRAY22 = BW2/2, var_type = "float";
assign BRAY23 = BW2/2, var_type = "float";
assign BRAY24 = -(BW2/2), var_type = "float";
assign BRAZ21 = -SPZ22, var_type = "float";
assign BRAZ22 = -SPZ22, var_type = "float";
assign BRAZ23 = -SPZ22+EPL22, var_type = "float";
assign BRAZ24 = -SPZ22+EPL22, var_type = "float";
assign BRAX25 = PX2+EPT2, var_type = "float";
assign BRAX26 = PX2+EPT2, var_type = "float";
assign BRAX27 = PX2+1.75*EPL22, var_type = "float";
assign BRAY25 = -STt2/2, var_type = "float";
assign BRAY26 = -STt2/2, var_type = "float";
assign BRAY27 = -STt2/2, var_type = "float";
assign BRAZ25 = -SPZ22, var_type = "float";
assign BRAZ26 = -SPZ22+EPL22, var_type = "float";
assign BRAZ27 = -SPZ22+EPL22, var_type = "float";
assign BRA_STIFF_X2 = PX2+1.75*EPL22, var_type = "float";
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
vert1 = Px1, Py1, Pzt1,
vert2 = Px1, Py1, Pzb1,
vert3 = Px1, -Py1, Pzb1,
vert4 = Px1, -Py1, Pzt1,
class = 1, grade = "A36", material = "Steel", name = "EP_0.028", 
thickness = EPT1;
plc_area
vert1 = -Px2, Py2, Pzt2,
vert2 = -Px2, Py2, Pzb2,
vert3 = -Px2, -Py2, Pzb2,
vert4 = -Px2, -Py2, Pzt2,
class = 1, grade = "A36", material = "Steel", name = "EP_0.019", 
thickness = EPT2;
plc_area
vert1 = spx11, spy11, 0,
vert2 = -spx11, spy11, 0,
vert3 = -spx11, spy12, 0,
vert4 = spx11, spy12, 0,
class = 1, grade = "A36", material = "Steel", name = "SP_0.012", 
thickness = STt1;
plc_area
vert1 = spx11, -spy11, -STt1,
vert2 = -spx11, -spy11, -STt1,
vert3 = -spx11, -spy12, -STt1,
vert4 = spx11, -spy12, -STt1,
class = 1, grade = "A36", material = "Steel", name = "SP_0.012", 
thickness = STt1;
plc_area
vert1 = spx11, spy11, -SPZ11,
vert2 = -spx11, spy11, -SPZ11,
vert3 = -spx11, spy12, -SPZ11,
vert4 = spx11, spy12, -SPZ11,
class = 1, grade = "A36", material = "Steel", name = "SP_0.012", 
thickness = STt1;
plc_area
vert1 = spx11, -spy11, -SPZ12,
vert2 = -spx11, -spy11, -SPZ12,
vert3 = -spx11, -spy12, -SPZ12,
vert4 = spx11, -spy12, -SPZ12,
class = 1, grade = "A36", material = "Steel", name = "SP_0.012", 
thickness = STt1;
plc_area
vert1 = BRAx11, BRAy11, BRAZ11,
vert2 = BRAx12, BRAy12, BRAZ12,
vert3 = BRAx13, BRAy13, BRAZ13,
vert4 = BRAx14, BRAy14, BRAZ14,
class = 1, grade = "A36", material = "Steel", name = "SP_0.012", 
thickness = STt1;
plc_area
vert1 = BRAx15, BRAY15, BRAZ14,
vert2 = BRAx16, BRAY16, BRAZ15,
vert3 = BRAx17, BRAY17, BRAZ16,
class = 1, grade = "A36", material = "Steel", name = "SP_0.012", 
thickness = STt1;
plc_area
vert1 = BRA_STIFF_X1-STT1, BRA_STIFF_Y11, BRA_STIFF_Z11,
vert2 = BRA_STIFF_X1-STT1, BRA_STIFF_Y12, BRA_STIFF_Z12,
vert3 = BRA_STIFF_X1-STT1, BRA_STIFF_Y13, BRA_STIFF_Z13,
vert4 = BRA_STIFF_X1-STT1, BRA_STIFF_Y14, BRA_STIFF_Z14,
class = 1, grade = "A36", material = "Steel", name = "SP_0.012", 
thickness = STt1;
plc_area
vert1 = BRA_STIFF_X1, BRA_STIFF_Y15, BRA_STIFF_Z11,
vert2 = BRA_STIFF_X1, BRA_STIFF_Y16, BRA_STIFF_Z12,
vert3 = BRA_STIFF_X1, BRA_STIFF_Y17, BRA_STIFF_Z13,
vert4 = BRA_STIFF_X1, BRA_STIFF_Y18, BRA_STIFF_Z14,
class = 1, grade = "A36", material = "Steel", name = "SP_0.012", 
thickness = STt1;
plc_area
vert1 = -BRAx21, BRAy21, BRAZ21,
vert2 = -BRAx22, BRAy22, BRAZ22,
vert3 = -BRAx23, BRAy23, BRAZ23,
vert4 = -BRAx24, BRAy24, BRAZ24,
class = 1, grade = "A36", material = "Steel", name = "SP_0.009", 
thickness = STt2;
plc_area
vert1 = -BRAx25, BRAY25, BRAZ24,
vert2 = -BRAx26, BRAY26, BRAZ25,
vert3 = -BRAx27, BRAY27, BRAZ26,
class = 1, grade = "A36", material = "Steel", name = "SP_0.009", 
thickness = STt2;
plc_area
vert1 = -BRA_STIFF_X2, BRA_STIFF_Y21, BRA_STIFF_Z21,
vert2 = -BRA_STIFF_X2, BRA_STIFF_Y22, BRA_STIFF_Z22,
vert3 = -BRA_STIFF_X2, BRA_STIFF_Y23, BRA_STIFF_Z23,
vert4 = -BRA_STIFF_X2, BRA_STIFF_Y24, BRA_STIFF_Z24,
class = 1, grade = "A36", material = "Steel", name = "SP_0.009", 
thickness = STt2;
plc_area
vert1 = -(BRA_STIFF_X2-STT2), BRA_STIFF_Y25, BRA_STIFF_Z21,
vert2 = -(BRA_STIFF_X2-STT2), BRA_STIFF_Y26, BRA_STIFF_Z22,
vert3 = -(BRA_STIFF_X2-STT2), BRA_STIFF_Y27, BRA_STIFF_Z23,
vert4 = -(BRA_STIFF_X2-STT2), BRA_STIFF_Y28, BRA_STIFF_Z24,
class = 1, grade = "A36", material = "Steel", name = "SP_0.009", 
thickness = STt2;
plc_area
vert1 = spx11, spy11, -SPZ21,
vert2 = -spx11, spy11, -SPZ21,
vert3 = -spx11, spy12, -SPZ21,
vert4 = spx11, spy12, -SPZ21,
class = 1, grade = "A36", material = "Steel", name = "SP_0.009", 
thickness = STt2;
plc_area
vert1 = spx11, -spy11, -SPZ22,
vert2 = -spx11, -spy11, -SPZ22,
vert3 = -spx11, -spy12, -SPZ22,
vert4 = spx11, -spy12, -SPZ22,
class = 1, grade = "A36", material = "Steel", name = "SP_0.009", 
thickness = STt2;
plc_area
vert1 = -0.118, -0.033,-0.068,
vert2 = -0.118, -0.033,-0.082,
vert3 = -0.118, -0.045,-0.089,
vert4 = -0.118, -0.057,-0.082,
vert5 = -0.118, -0.057,-0.068,
vert6 = -0.118, -0.045,-0.061,
class = 1, grade = "A36", material = "Steel", name = "HTB_M16", 
thickness = 0.013;
plc_area
vert1 = -0.118, 0.057,-0.068,
vert2 = -0.118, 0.057,-0.082,
vert3 = -0.118, 0.045,-0.089,
vert4 = -0.118, 0.033,-0.082,
vert5 = -0.118, 0.033,-0.068,
vert6 = -0.118, 0.045,-0.061,
class = 1, grade = "A36", material = "Steel", name = "HTB_M16", 
thickness = 0.013;
plc_area
vert1 = -0.118, -0.033,-0.128,
vert2 = -0.118, -0.033,-0.142,
vert3 = -0.118, -0.045,-0.149,
vert4 = -0.118, -0.057,-0.142,
vert5 = -0.118, -0.057,-0.128,
vert6 = -0.118, -0.045,-0.121,
class = 1, grade = "A36", material = "Steel", name = "HTB_M16", 
thickness = 0.013;
plc_area
vert1 = -0.118, 0.057,-0.128,
vert2 = -0.118, 0.057,-0.142,
vert3 = -0.118, 0.045,-0.149,
vert4 = -0.118, 0.033,-0.142,
vert5 = -0.118, 0.033,-0.128,
vert6 = -0.118, 0.045,-0.121,
class = 1, grade = "A36", material = "Steel", name = "HTB_M16", 
thickness = 0.013;
plc_area
vert1 = -0.118, -0.033,-0.298,
vert2 = -0.118, -0.033,-0.312,
vert3 = -0.118, -0.045,-0.319,
vert4 = -0.118, -0.057,-0.312,
vert5 = -0.118, -0.057,-0.298,
vert6 = -0.118, -0.045,-0.291,
class = 1, grade = "A36", material = "Steel", name = "HTB_M16", 
thickness = 0.013;
plc_area
vert1 = -0.118, 0.057,-0.298,
vert2 = -0.118, 0.057,-0.312,
vert3 = -0.118, 0.045,-0.319,
vert4 = -0.118, 0.033,-0.312,
vert5 = -0.118, 0.033,-0.298,
vert6 = -0.118, 0.045,-0.291,
class = 1, grade = "A36", material = "Steel", name = "HTB_M16", 
thickness = 0.013;
plc_area
vert1 = 0.143, -0.060,-0.066,
vert2 = 0.143, -0.060,-0.084,
vert3 = 0.143, -0.075,-0.092,
vert4 = 0.143, -0.090,-0.084,
vert5 = 0.143, -0.090,-0.066,
vert6 = 0.143, -0.075,-0.058,
class = 1, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.143, 0.090,-0.066,
vert2 = 0.143, 0.090,-0.084,
vert3 = 0.143, 0.075,-0.092,
vert4 = 0.143, 0.060,-0.084,
vert5 = 0.143, 0.060,-0.066,
vert6 = 0.143, 0.075,-0.058,
class = 1, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.143, -0.060,-0.141,
vert2 = 0.143, -0.060,-0.159,
vert3 = 0.143, -0.075,-0.167,
vert4 = 0.143, -0.090,-0.159,
vert5 = 0.143, -0.090,-0.141,
vert6 = 0.143, -0.075,-0.133,
class = 1, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.143, 0.090,-0.141,
vert2 = 0.143, 0.090,-0.159,
vert3 = 0.143, 0.075,-0.167,
vert4 = 0.143, 0.060,-0.159,
vert5 = 0.143, 0.060,-0.141,
vert6 = 0.143, 0.075,-0.133,
class = 1, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.143, -0.060,-0.216,
vert2 = 0.143, -0.060,-0.234,
vert3 = 0.143, -0.075,-0.242,
vert4 = 0.143, -0.090,-0.234,
vert5 = 0.143, -0.090,-0.216,
vert6 = 0.143, -0.075,-0.208,
class = 1, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.143, 0.090,-0.216,
vert2 = 0.143, 0.090,-0.234,
vert3 = 0.143, 0.075,-0.242,
vert4 = 0.143, 0.060,-0.234,
vert5 = 0.143, 0.060,-0.216,
vert6 = 0.143, 0.075,-0.208,
class = 1, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.143, -0.060,-0.291,
vert2 = 0.143, -0.060,-0.309,
vert3 = 0.143, -0.075,-0.317,
vert4 = 0.143, -0.090,-0.309,
vert5 = 0.143, -0.090,-0.291,
vert6 = 0.143, -0.075,-0.283,
class = 1, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.143, 0.090,-0.291,
vert2 = 0.143, 0.090,-0.309,
vert3 = 0.143, 0.075,-0.317,
vert4 = 0.143, 0.060,-0.309,
vert5 = 0.143, 0.060,-0.291,
vert6 = 0.143, 0.075,-0.283,
class = 1, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.143, -0.060,-0.366,
vert2 = 0.143, -0.060,-0.384,
vert3 = 0.143, -0.075,-0.392,
vert4 = 0.143, -0.090,-0.384,
vert5 = 0.143, -0.090,-0.366,
vert6 = 0.143, -0.075,-0.358,
class = 1, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.143, 0.090,-0.366,
vert2 = 0.143, 0.090,-0.384,
vert3 = 0.143, 0.075,-0.392,
vert4 = 0.143, 0.060,-0.384,
vert5 = 0.143, 0.060,-0.366,
vert6 = 0.143, 0.075,-0.358,
class = 1, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.143, -0.060,-0.751,
vert2 = 0.143, -0.060,-0.769,
vert3 = 0.143, -0.075,-0.777,
vert4 = 0.143, -0.090,-0.769,
vert5 = 0.143, -0.090,-0.751,
vert6 = 0.143, -0.075,-0.743,
class = 1, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.143, 0.090,-0.751,
vert2 = 0.143, 0.090,-0.769,
vert3 = 0.143, 0.075,-0.777,
vert4 = 0.143, 0.060,-0.769,
vert5 = 0.143, 0.060,-0.751,
vert6 = 0.143, 0.075,-0.743,
class = 1, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.143, -0.060,-0.826,
vert2 = 0.143, -0.060,-0.844,
vert3 = 0.143, -0.075,-0.852,
vert4 = 0.143, -0.090,-0.844,
vert5 = 0.143, -0.090,-0.826,
vert6 = 0.143, -0.075,-0.818,
class = 1, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
plc_area
vert1 = 0.143, 0.090,-0.826,
vert2 = 0.143, 0.090,-0.844,
vert3 = 0.143, 0.075,-0.852,
vert4 = 0.143, 0.060,-0.844,
vert5 = 0.143, 0.060,-0.826,
vert6 = 0.143, 0.075,-0.818,
class = 1, grade = "A36", material = "Steel", name = "HTB_M20", 
thickness = 0.016;
