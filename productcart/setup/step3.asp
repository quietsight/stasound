<%@ LANGUAGE = VBScript.Encode %>
<%#@~^tAAAAA==@#@&@#@&"+kwW	/ 2XwrD/C8kWV!O+,'PgGA`*P Pq@#@&]/2W	d+cb[[_+l[nMPEwMCT:Cr~rxW mm^tJ@#@&"+d2Kx/ b9NCC[+MPrmCm4n mGxD.W^JSEaDk-CD+E@#@&]/2Kxk+R;l14+;WUYMWs~{PJ	GO1lm4nE@#@&@#@&xzUAAA==^#~@%>
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/rc4.asp"-->
<!--#include file="pcXMLCheck.asp"-->
<!--#include file="../includes/contactEmail.inc"-->
<%#@~^7DUAAA==~7l.G?g'//krW	`EYsw/G	x+1OkKx?D.rxTJ*@#@&@#@&(6PHr:~\mD9jg'JE~:tnx@#@&i?nk/bWxvJa^\|fj1r#x-mDfUHP@#@&2	[~q6@#@&@#@&B~#mV;+,WWMPO4PE.s,4nk	L,Dn$E/YN@#@&q6P.+$EndDR0K.:vJ?!8hkDsKDhJb@!@*JEPD4+	@#@&7i@#@&7[b:~k	O;WhsE	kmmYbGxAD.@#@&drUDZWshE	kmmOrW	2MDx!@#@&i@#@&dbWPk+ddbWxcEams{Aa1+n9+9SkskDE#{Je2UJ~O4+x@#@&di/YM3UYMXAD.W.xr@!Ol(s+@*@!&OM@*@!O[@*K4+,U!:8D,W0,lDO+swO/,YG~M+obdYDPh.GN!mDZCDO~SkOt,rx1W..mY~^M+[+	OblskP6m+9dPDtnPslarsE:,UEs4+M~CV^WS+[R@!8M@*~n^nlk+~^^W/n~DtnP(.KhdD,lx9PD.X,lLlbx @!JYm@*@!zDD@*@!&Ol(V@*E@#@&7iDn/aGxk+ .Nk.n1Y~JkOwfclkwgk+D;w{trN[hdT'JLj+M\+M iId2	mGNnckY.2	ODH2..KD#@#@&i+UN,r6@#@&iBJzPV+D~0KDhP7l.rm4VdPmxN,dnY,k	Pd+ddbWU@#@&7Nb:~21\{3hmksnm.DxnM@#@&dNb:,2m7{F+Hq9@#@&dNbhPam\|9j1@#@&iNr:~21\mf~PXa+@#@&iNkh~am-{UOKDn`Id@#@&iNbhPam-{UYG.jqG@#@&iNks~2m7{UYGDnKqf@#@&i@#@&iw^-|2:Cr^nCDDUDxDDb:`M+$;+kY 0KDhcr2:mrVhlDDUnDr#*@#@&djnk/rW	cJam-mA:lrshl.Y	nMJb{w1\{A:mrVhl.Y	+.@#@&dw1-{n+X&9xYMks`.+5;/OR6GDs`EFXq9E*#@#@&ij/dbW	`Jam7m|X(fr#x21\{nnX&f@#@&72m7{G?H'O.b:cD5E/O 6WDhcrfj1rb*@#@&i?//bW	cJam-{G?HE*'w1-{G?1@#@&7w1\|f$Kz2'ODbh`M+5;/Y WKDh`r9~Kza+r##@#@&ij+k/rW	`E21\{G$KHw+rbxw1\|f$Kz2@#@&da^\|?OGM+j]J{Y.kscM+5!+kYR6WMh`r?OWM+i]dJ#*@#@&i?+kdrW	`r?OW.n`IJJ*xw1\mjDWDni"S@#@&i21\mUYKD+`qGxYMkh`M+5;/YcWWM:`rjOWM+`q9Jbb@#@&7?d/bWUcrwm-mUYGDi&fE*'am\|?DGDj(f@#@&721\{UOWM+nq9xYMks`.+5;/OR6GDs`EjDWDnKqfE#*@#@&dj/kkW	`r2m7{jYKDnKqfJ*xw1\{UOGDnqf@#@&7@#@&dvzJ~Mxn.mY+~3MDGD,\/dmoP0KD,n:aYzP6kns9/@#@&7Nb:PkO.2	YMX3D.GM@#@&d@#@&dkY.3	YDz3MDGD{Er@#@&iq6Pw1\|3:mksnmDOUD'rEPDt+	@#@&dik6PdY.3	Y.XA.DKD@!@*rJPO4x~/D.AxOMXADDKD{dYM2UYMX3.MWDLE@!(DPJ@*E@#@&di/OD3UDDz2M.WM'dOM2xO.H2.DK.LJ3 :mkVJnm.Y	+.P&f~rkPl,.+$EkMn[P6kV[RE@#@&d3x9~q6@#@&7&0P2^7{F+H(G'ErPDt+	@#@&7db0~/DD3UDDXA.DKD@!@*EEPDtx~/O.AxODH3DMW.xkYD3UDDz2M.KD'r@!(DPJ@*r@#@&iddYM2UOMX2M.WM'/D.3xDDH2.DG.LJF+H~q9Prd,lP.n$ErD[,0rV9RJ@#@&i3x9P(0@#@&7(6Pw1-{G?1{EEPDtx@#@&77b0~/D.2	Y.zADDG.@!@*EJ,O4+U,/DD2	YMz2MDGD{/O.AxYMz2MDWM'E@!(D,z@*J@#@&iddYM3xDDz3MDW.xkY.2	OMX3MDKD[rZKUxmOkKx~jDDk	LPb/Pm~.+$EbDnN~Wb+sNcE@#@&d3U9PqW@#@&d(0,21\mUYKD+`IdxJrPOtx@#@&idk6~/DD2	O.XADMW.@!@*ErPOtUPkY.3	YDz3MDGD{dDD3	YMX2MDK.[r@!8D,z@*E@#@&didYM2xD.z2MDKDx/O.AxODH3DMW.'r?YG.PiId~b/~mPM+;!kMnN,0r+^N E@#@&dAUN,q0@#@&7q6Pam-{jOKDnj&9'rJ~O4+x@#@&idr0,dDD3	YMX2MDK.@!@*JEPDtnU,/YM3xDDXA..WM'kY.2UOMX3DMGDLJ@!8MPz@*E@#@&7dkOM2UDDH2DMWMx/DD3xDDz3MDWM'J`/+M~(f,kkPCP.n$ErD[P6kns9RJ@#@&i2UN,(6@#@&iq6Pw1\|jYKDnnqfxErPY4nx@#@&dirWPkYM2UY.zAD.WM@!@*rJ~O4+x~dDD3xD.H2.MWM'/DDAUYMX3DMW.'r@!4M~z@*J@#@&77/DDAxODz3MDGD{dYM2UOMX2..KD'JhCk/AKD9PkkPm~D;;kM+[~6k+^[Rr@#@&i3UN,q6@#@&dv#mVrNmO+,srn^NP/G	YnxD@#@&d(6PgrK,k	dYM`2m7{jOKD+`]SBJtDO2)Jzr#xF~O4+U@#@&7db0~dDD2UOMX3DMGM@!@*rJ,Ytx,dYM2UYMX3.MWD{dYM2xD.z2MDKD'J@!8MP&@*r@#@&iddOM2xO.H2.DK.{/OM2	YDH2M.WM[E5KE.~UYWMnBkPb9[.+k/,:;/O~1WUYmrx,JE4DYwl&JJEPDG,4n,l,\l^k9~j"S J@#@&73	NP&W@#@&dq6~2m7{UYGDni&f@!@*rEPDtnU@#@&d7(6PHr:~b/H!:Dk1`a^\|?OWM+i(G#PD4+	@#@&i77k6PkY.2UOMX3DMGD@!@*EE,YtnU,/ODAUDDzADMWD{/D.2	Y.XAD.GM[J@!8D,z@*r@#@&didkY.2UOMX3DMGD{/O.AxY.zAD.WM'rjdD,qf,kk~xKY~\mVr[cJ@#@&7dAxN,(W@#@&dAx[PrW@#@&7B&WPDtn.Pl.n,lUX,nMDGM/,lY,Y4r/,wGk	YS~kYWa~l	NPMn[kM+1Y@#@&7r6P7/D.2	Y.zADDG.@!@*EJ,O4+U@#@&idD/aGxk+ DNr.mY,E/D+w2 C/agk+OE2xDD;+Lh/T'E'U+D-nMRiId3	mG9+v/YM2	ODH2.DKDb@#@&d+	[Pb0@#@&@#@&db0,/^fjH@!@*EJ,Otx@#@&idB&&Jz&zJ&Jz&JzJzzJzJ&zJz&zJz&&JzzJ&zJzzJ&&zJzJz&z&&Jz&zJ&zJz&&Jzz&&Jz@#@&i7Ez&,K4k/,/DGDP4lkP2.\kK;/^XP(nnx,Dor/OnM+[@#@&7dEdq ,Ztn^0POtmO,+ab/DkxTPk^fU1~/DDrUTPkk~\mVk9@#@&diBi  P/4mVPD4lDPsGTkx~^M+[+	OblskP6WD,Y4nPmNhk	PhCDmt,AkDtPD4nP94@#@&7dv72R~ZMnlD+~UhPdOKDnZKUkYC	YkRlkw,Wk^+@#@&idv&JzzJ&zJzzJ&&zJzJz&z&&Jz&zJ&zJz&&Jzz&&Jz&zJ&Jz&JzJzzJzJ&zJz&zJz&&Jzz@#@&di@#@&i7@#@&idEz&Pq ,Z4+1VPDtCO,+6rdDkUo,d1fjgPkYDbxT~kkP-l^k[@#@&dd9r:,mW	UP+swBP5En.H~~Dk@#@&iddnDPmGU	Kn:axk+.7+MRmM+mO+K4%+1YcEmNW98R1Wx	n^YbW	Jb@#@&7iB6wUPHW;.,mWUUmOkKU,YG,Y/Y,k6~mKxU+1YrG	P/D.k	oPbd~\mVbN@#@&771WUx:n:aR62xPd^G?HP@#@&idr6PDDcx!h4D@!@*ZPO4x@#@&7di+DM UEs4Dx!@#@&id7mKUx:+h2crwnU,Y.kscM+5!+kYR6WMh`rfj1r#b@#@&ddir0,+DM UEs4D@!@*T~Dtnx@#@&did7^KxxPnsw m^Gk+@#@&diddk+D~mKxUK:2~{PxKOtbxo@#@&7didEmGxU~kY.k	LPb/~UKYPC~7lsk9~kY.bxT@#@&idi7D/2W	/n M+Nb.+1YPrdO+a&cldw_dY;w{OD!+'hko'E'k+.\.cj]d2	mW9+vEqDPCwa+C.kPY4CY,XW!.~mmDDP4ld~(+nx,2D\rG!/Vz~M+LkkODn9PSkY4Pm~mKxU+1YrG	P/D.k	oPD4CY,kkPUWO~7lsk9 P"+LrkYDCObWUP1C	xGDP1WxDk	;+cJb@#@&d77iD+k2W	/+c3UNv#@#@&7d7n	N~k6@#@&idnU9PkW@#@&d7@#@&7iB&JPg2,O,)N9P^tmV~DWPkn+,k0,O4kkP1l.Y~rkPCV^GhN~OKP4n~M+LkkODn9Pmolbxc@#@&idnDMR^slD@#@&di+DM UEs4Dx!@#@&id9ks~W(Lp\dCKPKBPa:^@#@&@#@&idE?+	N,OtP.+$EndD@#@&i7/D+6DxEw1.DdkGU{J'/1#+M/rG	@#@&77kYn6DxkYnXY,[Pr[a^?!4#+M/rG	'J,'Pkm?!8#+M/bWU@#@&7i/O+XO'kYnaDP[~ELw^|z&NxrPLPw1\|F+Hq9@#@&d7dD+6Dx/D+6D~'Pr[am9APza+xJ,'Pam-mGAKz2@#@&didD+aD'kY+XY,'Pr[2mA:Cr^nlMOxD'r~'Pam7{3:Cr^nCDDU+M@#@&7i/YnaD'dYaDP',JLwmUYK.+`IJ'rP'~am\|jYKD+`]J@#@&di/O+aO{/O+XOPLPE'am?OGM+KGxrP',w1\{UYK.+h9@#@&d7dD+6Dx/D+6D~'Pr[amjYG.j(f{EPLP2^7{?OGM+iqG@#@&d7kY6Y{/Dn6DP'Pr[2^U+/krW	qf{E~[,//dkGUc?n/krW	q9@#@&dddO6O'kO6O,[,J[am"nob/O+M+[H!:4.'rP[,d^Iob/O+.n9@#@&didY6OxkY+aO,[~JL21ZGswmxXglsn'rP'Pkm/Gswl	z1m:+@#@&7dkY6O'dO6OPL~JLw^/K:wCUHb[NMnk/xrPLP/1ZKhwmxzb9N.nk/@#@&7dkY+XOx/D+XY~[~ELw^ZKhwmxztbw'E~LPdm;GswC	X\kw@#@&i7/D+aY{/OnXYPL~JLwm;GhwmxHZrYzxrP'Pk^ZK:2C	XZrOH@#@&didD+aD'kY+XY,'Pr[2m;Wh2mxXUOlD+'r~'Pkm;WhwCUH?OlDn@#@&d7dD+6OxkYn6D~LPELw1ZWswmUX;W;xDDzxrP[,dm;W:aCUX;W!xODz@#@&d7/Dn6D'dO6Y~',J'w1)^+.DKHw+{fUH2oqjKUJ@#@&id/Dn6D'/DnaY,[,J'DnW{/nY!2J@#@&@#@&idBjn	N~Y4n,Y.mxklmDkKUPbxWW,ld~alDD~W6PY4n~;!+MXdY.r	o@#@&i7/Y~asVPx~U+.\.cZ.lD+r(L^YvJ\/X:s+c/+M-+M(:^uOYaJL6j+O#Db@#@&7dX:s Kw+U~rn6?:EBPE4YDw)JzSAhc+CD^XrhalmD mK:za.GN!mDmCDO&7+.k6zzamFnH.+.r6XiId m/2ygr[PkYaY,[~Jr~~WmV/@#@&id6ss /x9PEJ@#@&@#@&7dbUY;Whh!xk^CDkGxA.M'T@#@&idBJz,/tmVP6W.~1Wx	nmDkW	~r/kE/@#@&77b0~+M.R	Eh8D@!@*T,Y4+	@#@&d7iB:t+M+,r/,l~mK:h;	kmmOkKxPbddEP P]+LrkYnD,OtP^CMYPCd,kdR@#@&id7bxDZWs:!Uk1lOkKx3.M'F@#@&did?O~6sV,'~1GO4kUo@#@&di+U[,k0@#@&id@#@&i7b0~bxDZWs:!Uk1lOkKx3.M'!,Otx@#@&77dkYM?OlO;kPxPXhVc?OCDE/@#@&id7@#@&7idr6PkYDUYmOEk@!@* Z!~O4+x@#@&diddEP4+M+,kdPC~1Wh:!Uk1lOrKxPrdkEnP ~"+Lb/D+D,Y4nP1l.Y,ld~b/R@#@&diddbUOZK:sEUk^CDkGxA.D{F@#@&idd7jY~6ss,'~gWDtk	o@#@&didnVk+@#@&iddiv/DWD~OtPM+dwGUk+@#@&i7di/O."+Y#C^PxPXh^R./aWxk+:n6D@#@&did7jYPXhV,'PgGOtbxT@#@&d77i/ODz.DmX#C^P'~daVrYvdDD]YjlVBPrkJBPRF*@#@&7idda^\|?YmO;/,',/OD).Mlz.ms`Z#@#@&iddnU9Pr0@#@&idn^/@#@&idi2m7{jYmY;d{JfAob`SKr@#@&di+	N~kW@#@&d7@#@&7dEz&~b0Po)&S~D^KD[,///bW	~mKEUY@#@&77b0Pa^\|?YmO;/{Jwb(SE~Dtnx@#@&didrW,/+ddbWU`r21\mU+k/kKx;UYr#xJrPO4x@#@&7did/ddkKxvJ2m-mU+d/bGx;xOE*'F@#@&id7+^d@#@&idid//krW	`Ew1\mj//bGx;xYrbx//kkGxcEam-{Un/kkGU;xYEb3F@#@&i7i+U9Pb0@#@&di7@#@&d7db0~d//bGxvJw1-m?/kkGx/UDJb@*l~Y4+U@#@&dd7@#@&d7divJz~U+	NPmV.Y,YGPAq~-blP_PKh@#@&i77dkY6O'E21{#+MdkKxxfc!*q'9+LDn{FE@#@&iddi/Dn6D'dY6O~LPJL2mA:lbsKlMY	+.'E~LP2m7m2slrshlDOUD@#@&i7iddD+XY'kYaY,[~JLw^FXq9xJ,[Pa^-{n+Hq9@#@&7id7/Dn6D'dO6Y~',J'w19U1xrPLPw1\|9?g@#@&did7dD+6Dx/D+6D~'Pr[am9APza+xJ,'Pam-mGAKz2@#@&di7i/O6D'/D+XOPLPE[amjOKD+`]S{JPL~2m7{UYGDni"S@#@&i7di/OnXY'dO6OPL~r[21?DWDnq9'rP'Pam-mUYWMnnqf@#@&77di/D+aYxdD+aY,'Pr[2^UYW.n`q9'r~LP21\|?YKDiqG@#@&did7dD+6Dx/D+6D~'Pr[amj+ddbWUqGxJ,[~d//rG	Rj+kdbWU&f@#@&di@#@&7didvbd2]P,2q,6s,nrUj(Ad2,s]bi9@#@&7di7/Y~asVPx~U+.\.cZ.lD+r(L^YvJ\/X:s+c/+M-+M(:^uOYaJL6j+O#Db@#@&7didah^RW2n	PEn}j:JS,J4YYa)J&hSh +mDszb:wm^YcmWs&2DKN!mOmC.Dz-+Mr0Hz2^n+X)sDORmdagELPkY+XY,'PrJSP6lsd@#@&i7di@#@&i77dX:^Rd+U[,JE@#@&7diddOM?YCO!/~',asV UYmYEk@#@&7didr0,/O.UYlD;/@!@* ZT~Y4+	@#@&d77idvK4nDPrd,lP^Gs:;xb^mYrKx,k/kE~O,Inob/OnMPY4nP1lDD~C/,kkR@#@&77id7k	OZK:h;	kmCObWU2M.{F@#@&diddi?OPX:sP{PHGDtk	L@#@&ddi7nx9Pb0@#@&77idj+D~6sV~x,1WO4bxL@#@&7id7@#@&iddiBJ&PzVnDDP^;kYWsnD,YtmO~Y4+HPhEdO,msWknP9WAU,YtnrMP8DKAk+.@#@&iddi/D.2	Y.XAD.GM'J@!Ol(V+@*@!&YM@*@!Y[@*P4PUEs8+MPGW,lYOnswO/,OKP.ob/YD,KDKN;mDZC.DPhbOt,kx1G.DmDP^Dn[xOkms/,+a^+Nd~DtnPsCXkh!:,xEs4.PmVsWS+[ @!4D@*~n^+lkn~m^Wk+~Y4n,4.WSd+MPCU9PY.z,lLlbUc@!&Dm@*@!zDD@*@!zDl8V@*E@#@&ddi7D/wKUd+cDNrDn^DPE/Dnw2RCdag/nO!wxtb[[hko{J[U+M-+MRiId2U^KN+vdYM2xD.z2MDKDb@#@&7id7@#@&7di+U[,k0@#@&id7/D.AxOMXADDKD{dYMb.DmX#C^`&*@#@&iddMndwKxk+ Dn[bDnmD~JkYn22Rld2Q/nY!2{Y.!+L:/T'r'?D-+MRi]d2x1GN`/D.3xDDH2.DG.*@#@&dinx9PrW@#@&@#@&7iB&z,+cP/4+13PDtmOP^WLk	P^.N+	OkmV/,WGD,Y4+~l[hbx~:mOm4PArDtPO4P[4,@#@&d7b0,+DMR	;:(+.'ZPO4x@#@&7dik0,O.ks`M+5EndDRWWMh`r?OGM+j(9r#b'rE,r],YMk:vD5E/OR6W.hvJ|zqGJ#*xEJ,Y4+U@#@&7id7DdwKxdncD+[rM+^Y,EkYna&cl/agknY!wxYMEn's/o{E[k+D7n.R`Id2UmG[`Eb^sP6kns9/PC.P.+$;bDn9Rr#@#@&di7+	N~k6@#@&7id/am|XqGxODb:vDn;;nkY 0K.:vJFnHqfEb*@#@&di7kwmmNskxalkdhKD['x9n;DXaO`DDksc.+$E/ORWGM:cJUOWM+K	GJ#bSkwm|z&fb@#@&idd$E.X{Jj2d2/P,kNm[:bxPw]6H,l9:rxdIr@#@&di7/Y~.k'mGU	Kn:a 6n1ED+`$E.X*@#@&didrW,1r:~DkR+KW~Y4+	@#@&d77i;;+Mz'r?3JAZK~r9l[:bU,s]}H,lNsk	dPqC3IAPr[mN:bU'r[YMrh`M+$En/O 6W.:vE?DW.n`qfEb*[EPzHGPC9:bxwm/kAWMNxBr[d2|lNsrxal/kAGD9[rBIJ@#@&id7dknY,Ddx1WxUP:2Ram;D+v;EDHb@#@&d7dikW~M/RG0,YtU@#@&idid7DndaWU/ DNr.mY~EkYnw2 m/2Q/YEa'D.E[h/T'E'k+D7nDcjId3UmKN`E5G;,NGP	GY,tC-Pw.Ga+.PMrTtOkPDWPME	~Y4kdP6ksncPK4r/,nDK[;mDZmDOP4CkP8+UPaDn-bWEdsHP.+TrkYnM+9Pl	N,OtPi/D~(GPNKn/,xWD~hlDm4POtn~`/nD,(f,WW~Dtkd~1l.Yrb@#@&7ididD/aGxk+ 2	Ncb@#@&di7dV/@#@&dididvz&~2R~ZMnlD+~UhPdOKDnZKUkYC	YkRlkw,Wk^+@#@&id77ik0,7/DD2	O.XADMW.@!@*ErPOtU@#@&d77idd.nkwGxkncDn9kM+mDPrdYwfRm/2_k+Y!2'DDE'h/T'r[j+.-D j"J2	mG[`/O.AxODH3MDGM#@#@&didi7+	N~k6@#@&7iddi@#@&iddi7d+k/bWU`EC9:rxrb'8@#@&@#@&dd77iDn/aG	/ncDNkM+1OPrR zbx^s!N+k&nmo+;.nlD+UYGDn/KxdYmUYkRCdag:G['qJ@#@&id7idM+/aW	d+c2UNv#@#@&iddinx9Pk6@#@&didVd+@#@&id7dE&z,& ~;D+COPU+S~kYGM+;WxkYmUYkRC/aPWr^+@#@&7didk6~7/DDAxODz3MDGD@!@*JrPO4x@#@&7id7dMnkwG	/RDNb.+1Y~JkYn22Rlk2gk+Y!2xYME[h/Lxr[j+M-+MRi]d2x^G9+c/D.AxOMXADDKD*@#@&id7dx[~b0@#@&7di@#@&i77dk+k/rWUcrl[:bUJ*'q@#@&dd77M+dwKUk+ M+9kDmD~JcR&k	ms;9+/JKlT+ZMnCY?DW.+/G	/Ol	O/cld2Q:W[n{FE@#@&7id7M+kwW	/ 2	Nc#@#@&77i+x9~k6@#@&i73x9P&0@#@&7n^/n@#@&7dEz&~2RP/.lO+,Uh~kYKD+;W	dYmxO/cld2,0k^n@#@&ddbW~dkYM2UY.zAD.WM@!@*rJ~O4+x@#@&id7DdaWUk+cD+9kMnmDPE/D+2fcl/a_/YEaxOD!+L:doxEL?nD7nDcj]JAxmG[`dYM3	Y.H2MDWM#@#@&di+UN,kW@#@&ddi@#@&idBJ&~1A,O~b[[,m4+1VPDW~d+PrW,Y4kk~1l.DPb/PmV^GhN~YKP8n,D+Tr/D+D[~lTlbx @#@&@#@&d7BUnx9PO4PDn5!+dY@#@&iddD+XY'rw1m.DdkKxxfc!*8E@#@&ddkOn6D'kYn6O~LPE[a^2slrshlDOUDxJ,',w^7{A:lbVhCDDxnD@#@&77kY+XO'kY+XO~[,JLw^|nz&NxJ,'Pam-mn+X(9@#@&7dkO6O{/D+6DPL~JLw^fU1xE,[Pa^\|f?g@#@&di/D+aYxdD+aY,'Pr[2^GAKz2'EPL~am-|f~KXa+@#@&di/O+XYxdD+6D~[,J[a^jYKDj]SxE,[~w1-{UYG.jIJ@#@&d7/DnXYxkY6Y,[,E[amjYKDnKqf'r~[,wm7mjYKDn	f@#@&iddYaY{/OnXYP'~r[2mUOKDn`qG'J,[,2m7{jYKDni&f@#@&7dkY+XOx/D+XY~[~ELw^?d/bWU(G'J~',/n/krKx U+k/kKx&9@#@&@#@&diBjn	NPD4+,YDmUdl1YbWUPrU6W~lk~wmDO~K0PO4P5E.H/OMk	o@#@&dinDMR^Vl.@#@&@#@&i7/YPXhsP{PU+.\n.cZ.+mO+}4%n1Y`E\k6hVy k+.7+M(:^CDOwr[a?Y#nM#@#@&7dX:VcG2+	Prn6?PEBPEtDOw=z&AShRnCMVzks2mmOcmK:zaDK[E1Y^lMY&-Dk6zzam|z#+Mk6X ld2QDn0{d+DE2'r[PdO6OPL~rJS,0mV/@#@&7dX:sRk+U[,JJ@#@&diBzJ~/tm0PWW.~1WUx^YbWU~b//;nk@#@&dir6PnMDcxEs4.@!@*!~Y4+U@#@&ddivK4+D~r/,l,mG:h;	k^lDrW	PrdkE+~R,InobdD+.,Y4+P1lMOPm/~kkR@#@&iddbUY;W:s;Uk1lDkGx3.M'q@#@&7di?nO,6:s~{PHWD4bxL@#@&id+	N,r0@#@&7d@#@&77b0PbUY;W:s;Uk1lDkGx3.M'TPD4+	@#@&7id/O.UYCY!d,'~X:^R?DlD;/@#@&7diddOMI+D#l^P',ahVcD/2WUdKn6D@#@&id7r6P/O.UYCY!d@!@*+Z!,Ytx@#@&did7B:tn.Pkk~l,mWsh;xbmmYrWU~b/dE~O,InLb/Yn.,Y4+,^mDO,lkPkkR@#@&did7k	Y/Gs:E	rmmYkKU3DM'8@#@&d77i?nY,a:^Px~gWY4r	o@#@&i7i+sk+@#@&didiv/DW.+,Y4n,D+k2W	/+@#@&7d@#@&id7djnDPa:^~',1GO4kxL@#@&d7didDD)MDmX.mV,xPkwskD`dOMI+D#l^~PrkE~,O8#@#@&77id2m7m?DlO;kP'~dDD)DMCH.C^`Z#@#@&di7+	N~k6@#@&7i+Vkn@#@&ddi2^\|?DlOEdxrf3sziS:J@#@&id+U[,kW@#@&@#@&d7EzJPk6Pw)qdP.+1W.[,/+kdkKxP1G;xD@#@&d7kW~am-{UOlDEdxrsb(JrPOtU@#@&7idb0Pk+kdkKxcJam-mU+/krW	ZxDEb'rJ,Y4+U@#@&d7did+k/rG	`J2^7{j+kdbWU;xDJ#{F@#@&didnVk+@#@&iddid+k/kKUcJam7{j+ddbWUZ	OJ*'dnk/kGUvJ2m7mU+dkkKxZ	Yrb_8@#@&didnU9Pk6@#@&idd@#@&7dik6Pd+ddbWU`r2m7{jnk/kGU;xOJ*@*lPO4+	@#@&idi@#@&id7dEz&~U+x9~l^+DD~OW,2&P-kC~_KPn@#@&did7dD+6Oxrw^{jnM/rKx{&RZ*8'No.+'qE@#@&di7dkY+XOx/D+XY~[~ELw^2sCk^nC.Dx+.xrP'Pa^7{3slbVnmDDU+M@#@&did7dD+6Dx/D+6D~'Pr[amF+z(9'EPL~w1\mFXq9@#@&d7didD+aD'kY+XY,'Pr[2mG?HxrP[,2m7{fUH@#@&ididdYnaD'dYaY,[~ELwm9$:X2+{E,[~am7{f~KH2+@#@&7diddO6Y{dY6Y,'~JLw1?OW.n`IJ'r~[,w^-|?YG.j]S@#@&id7i/D+6D'kO+XY~[,J'21?YK.+hf{E~[,w1\m?OGM+KG@#@&id77kY+aO{/O+XO,[~r[am?DWMnj&fxJ,[~21\{UOWM+j&9@#@&ididdYnaD'dYaY,[~ELwmjnk/rW	(G'E,[,/+k/bGxc?n/kkGU&f@#@&7d@#@&di77BU+	N~Y4n,Y.l	dl1YrG	PkUWKPC/,2mDO,W6PY4+,5EDz/DDrUT@#@&i7di/+D~a:^P{Pj+.-D ZMnlD+68N+mOcrHd6ssyRdD7+Do:^uYDwE[X?nOj+D*@#@&iddiahVcWa+UPEK}?PJB~J4YO2=zzAASRnlMsHkhal1YR1Ws&wMW[E1Y^CMYz7nDb0XJ2^|XzVnDO m/2gr'PkYnaDP[~Er~~0msk+@#@&didd@#@&i7di6hVc/nU9PJr@#@&iddidODUYmY;/~x,6hVcjYmY;d@#@&d77ikWPkOM?OmY!/@!@* ZTPDtnx@#@&77iddEPtD+,rdPmP1Wh:;UbmCYbGx,kdd!+PR~"+LkkOD~DtPmmDD~lkPr/c@#@&7iddirxDZWsh;xbmmYrWU3MDxF@#@&did77U+Y~asV~',HKY4bxT@#@&idi7+	N~k6@#@&7idd@#@&diddUnOPX:^PxPHGDtrxT@#@&id77@#@&d77iB&z,)^+.DP1E/DWsnD,Y4lDPO4XPs;/DPm^Gd+,NKhUPO4k.P(.WS/n.@#@&d77i/ODAUDDzADMWD{J@!Ol(Vn@*@!zO.@*@!Y9@*K4+P	;h4D,WWPCOD+hwDdPDW~.okdOD~nMG9E^DZmDY,hbOt,kUmKD.n1YP1.+9+xDrCVkP6^+n[kPOt~:m6rh!:PU;s4nD,C^VGS+9R@!(D@*~n^+C/P^sK/+,OtP4MGA/D,lUN~OMX~lTCk	R@!&Dm@*@!&DD@*@!JOm4s@*r@#@&idi7k6Pd+k/rG	`Ja^V|261nnNNdkhkOE*'EJ,Otx@#@&idd77k+d/bG	R)(l	NW	`*@#@&id7di/ndkkW	cJamV|3am+9+[SrhbYE#{E5A?E@#@&dd77x[PbW@#@&7idiD+kwKU/R.+9k.n1YPrdYw&cCdwQ/Y;wx4bNn[sdo{J'jD\n.cj]SAU1W[`kYDAxD.XAD.WM#@#@&iddi@#@&iddU[Pb0@#@&7d7@#@&d7dkODAxO.H2D.GM'dYM)MDCH.mV`2#@#@&did.+kwGUk+RMnNbD+1O~JkYwfRCdagd+D;w{Y.;[:dL{J'?.7+.cj"S2	mK[+v/ODAxO.H2DMGD*@#@&@#@&7dx9Pr0@#@&@#@&7dGr:,w^-|+rL4YixbO@#@&7iw1\{q+bLtDjUkD'.n$E+kO`r+bL4Y`xbYE#@#@&idr0,2m7{	nbotOi	kO'rE,Y4x@#@&dida^\|nkTtOi	kY{ES~?J@#@&7dx9Pr0@#@&idj+kdkKxcEam\m	kLtDi	kOr#{wm7{qnkTtOj	kO@#@&ddkn/kkW	cEl9:bxE#xq@#@&7@#@&7dM+d2Kx/n M+[kMn1Y~rRczk	m^;N/&nmon/M+lDn?DWD/GxkYmxO/ Ckw_:K[+{F'r	YInd{J'k	O;WhsE	kmmYbGxAD.@#@&d7./wKU/R2	[c#@#@&i+UN~r6@#@&+	[Pb0@#@&NukPAA==^#~@%>

<!--#include file="pcSetupHeader.asp"-->

	<script language="JavaScript">
	<!--
	function win(fileName)
		{
		myFloater=window.open('','myWindow','scrollbars=yes,status=no,width=620,height=400')
		myFloater.location.href=fileName;
		}
		
	function PCRegister_Validator(theForm)
		{
		if (theForm.StorePWD.value !== theForm.StorePWD2.value)
			{ 
			alert("The passwords do not match. Make sure you enter the same password in both the password and password confirmation fields.")
			theForm.StorePWD.focus()
			return (false);
			}
		return (true);
		}
	//-->
    </script>

		<form name="PCRegister" method="post" action="step3.asp" onSubmit="return PCRegister_Validator(this)" class="pcForms">
						<h1>Step 3: Activate your copy of ProductCart</h1>
						<%#@~^JwAAAA==~&0~M+$E+kYc}EDz?DDrUT`Jsdor#'rE~Y4+	P/wwAAA==^#~@%>
							<div class="pcCPmessageInfo"><u>All fields are required</u>. Enter the information exactly as it was provided to you when you purchased your license. <u>This server must be connected to the Internet</u> to activate your license.</div>
						<%#@~^BgAAAA==~VdP6QEAAA==^#~@%>
							<div class="pcCPmessage"><strong>There is a problem...</strong><br />
					  <%#@~^KgAAAA==~M+daW	/+chMrYP.+$EndDRp!nDH?YMrUovJs/LJbqA8AAA==^#~@%><br />For any activation issues, <a href="mailto:<%=#@~^DwAAAA==21\/KxDlmD2sCk^/QUAAA==^#~@%>">contact us</a>.</div>
						<%#@~^CAAAAA==~x[,k6PZgIAAA==^#~@%>
					  <%#@~^JAAAAA==~b0~M+$E+kYvE/Y;wr#@!@*rtk9nJ,YtU~JQsAAA==^#~@%>
						<table style="margin-top: 10px;">
							<tr>
								<td nowrap><div align="right">E-mail / Partner ID:</div></td>
								<td><input type="text" name="EmailPartner" size="45" maxlength="100" value="<%=#@~^GwAAAA==j/dbW	`Jam7m2slrVhl.O	+Drb5QkAAA==^#~@%>"></td>
							</tr>
							<tr>
								<td width="20%"><div align="right">Key ID:</div></td>
								<td width="80%"><input type="text" name="KeyID" size="45" maxlength="100" value="<%=#@~^FAAAAA==j/dbW	`Jam7m|X(fr#1wYAAA==^#~@%>"></td>
							</tr>
							<tr> 
								<td><div align="right">User ID:</div></td>
								<td><input type="text" name="StoreUID" size="40" maxlength="250" value="<%=#@~^FwAAAA==j/dbW	`Jam7m?DW.+`q9E*EAgAAA==^#~@%>"></td>
							</tr>
							<tr> 
								<td><div align="right">Password:</div></td>
								<td><input type="password" name="StorePWD" size="40" maxlength="250"></td>
							</tr>
							<tr>
								<td><div align="right">Confirm Password:</div></td>
								<td><input type="password" name="StorePWD2" size="40" maxlength="250"></td>
							</tr>
							<tr> 
								<td valign="top"><div align="right">Store URL:</div></td>
								<td><input type="text" name="StoreURL" size="70" maxlength="250" value="<%=#@~^EwAAAA==j/dbW	`JUYK.+`IJJ*eQYAAA==^#~@%>">
								<div style="padding-top: 5px; font-size:11px;">
								Enter the full URL to your store, before the &quot;productcart&quot; folder. For example, if your store is located at http://www.yourstore.com and the &quot;productcart&quot; folder is in the root of the site, you will enter the Store URL: &quot;http://www.yourstore.com/&quot;. If you rename the &quot;productcart&quot; folder, remember to edit the corresponding variable in the file &quot;includes/productcartfolder.asp&quot;.	</div>
								</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr> 
								<td><div align="right">Database Type:</div></td>
								<td>
									<select name="DBType">
										<option value="Access" selected>Access Database</option>
										<option value="SQL" <%#@~^MwAAAA==r6P(	?DD`dZmd+v?n/kkGUvJw1-{G?1rbb~r/$VE#@!@*ZPOtUiA8AAA==^#~@%>selected<%#@~^BgAAAA==n	N~b0JgIAAA==^#~@%>>SQL Server</option>
									</select>
								</td>
							</tr>
							<tr> 
								<td><div align="right">Connection String:</div></td>
								<td><input type="text" name="DSN" size="70" maxlength="250" value="<%=#@~^EgAAAA==j/dbW	`Jam7mfU1E#BgYAAA==^#~@%>"></td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr> 
								<td colspan="2">ProductCart supports both pounds and kilograms. This setting is NOT easily changed once the store is live. Please select the weight measuring unit that applies to your store.</font></td>
							</tr>
							<tr>
								<td><div align="right"><input name="WeightUnit" type="radio" value="LBS" checked class="clearBorder"></div></td>
								<td>Pounds</td>
							</tr>
							<tr>
								<td><div align="right"><input type="radio" name="WeightUnit" value="KGS" class="clearBorder"></div></td>
								<td>Kilograms</td>
							</tr>
							<tr>
								<td colspan="2"><hr></td>
							</tr>
							<tr> 
								<td colspan="2" align="center"><input type="submit" name="SubmitForm" value="Activate" class="submit2"></td>
							</tr>
						</table>
						<%#@~^CAAAAA==~x[,k6PZgIAAA==^#~@%>
			</form>