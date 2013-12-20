<%@ LANGUAGE = VBScript.Encode %>
<%#@~^vwEAAA==@#@&BP4kkP0bV~kkP2lMY~G6PnMGN!mY;C.YBPmx~+^Gs:nD1nPmw2sbmlOrKx~N-VGa+9Pl	N,dW^N~4HP3CMVX,(:almD~JS;R,n.W[;1Y/lMO~,kOd,/W;.1+~mK[~~DtPnMW9;mDZCDDPUCs+PmUN,VWTG~lM+,w.W2nMYzPKWPAl.sHPqh2mmO~,JdZ ,ZKwXMkT4Y, T!8O+TZ&R,)V^PDbL4YkPM+d+.-N PIGE,l.n,xWO~mVsWSn9POKP!/+BPmsYDSP9kdOMk4!O+,lx9&GD,D/nVs~mxzPaCDD/~G6Pn.G9E^Y;CMYvkPkWEMm~mKNnPSkO4KEY,OtPhMrOYx,mGxdn	Y~W6~2mDsz,q:2C1Y P:G,mG	YmmY,2m.VHP(:al^OBPw^nlk+P7rdkDPShARnCMVzks2l1Y ^K:R@#@&MpwAAA==^#~@%>
<%#@~^OgAAAA==~alLKbYV'rjtbw2k	o~	bylM[P PZmU^+^Pmx[P9n^+O+,jtbwhn	YJ~FBQAAA==^#~@%>
<%#@~^EgAAAA==~U+^DkKx'r:	Lb1mEP3wUAAA==^#~@%>
<%#@~^CQAAAA==Ksb[sk	',HAMAAA==^#~@%><!--#include file="adminv.asp"-->
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"-->
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/FedExconstants.asp"-->
<!--#include file="../includes/pcFedExClass.asp"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/productcartFolder.asp"-->

<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->  
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<%#@~^eBsAAA==~@#@&9b:,;EDHSPM/SP1WUUD+:a@#@&Gk:,rKlT+;E.DnUD~~\m.s^lL(	mWh2^+O+B~!+.H~,/YMr"9~,w^\|kUO}DN.qG@#@&GrhPam7{dY.\Y4W9Hls+S~am\mdDD\+D4KN]w^X~,Z!dYK:nD:DCUklmDrW	qNUOk6kDSP2^7{dYM)m1W;UD1Eh8DSPa^7{dDDt+YDg;:(+.~,w^-|/YM/lMDk./W9+@#@&9kh~am-{kOD:DC^0kxLH!:8+MS,w^7{kYDUtb2:xOb1mG;	Y1!h4D@#@&9r:,w1\m/O.G+dYbUlDkGU;WEUOMX/W9nBP21\|/YMfdYbxCYbWUKK/YmsZKN+B~2m7{kY.SCUTECo/W9+S~am\mdDDJW1C^+/KN~Pam7m/DD9+DlrsUml	d~,wm7mdYMnmorxLPK3nx@#@&fb:~WN+amaWdY9CDlS,W(LsNAaZ^ld/BPG8NrED2ED(Hd9GmBPkD-s39A(p:^uYDwS~w2f3p|Dn/!sD~~w2G2(|j"J~,w^\|/O.ADDK.Hko~,2^\|/DD)mOrKx@#@&@#@&@#@&@#@&v?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?@#@&EPjPzIKl~}1~S})G@#@&EU?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U@#@&@#@&vJz,M2:P}]fAI~qG@#@&21\{kOD}DN.(f{I;;+dOvJrNrb@#@&w^-|/Y.j/dkKU}D[D&f'U+kdkKxcJam)[skx}.NDqGEb@#@&k6P2m-mkY.?d/bWU6MN+.(G'EJ,6"Psxvwm7{kOD}D[+Mq9b@*!PD4+	@#@&i2^\|k	Y6D[nMq9'a^\|/O.}DNn.&f@#@&ij/dbW	`Jamz[:bx6D9+.(GJ#{2m7{k	O6D9+Mq9@#@&n^/n@#@&7w1\mr	Yr.[D(f{21\mkYM?+k/bGx}D[+Mq9@#@&+x9~k6@#@&@#@&vzJPhb!2~HzH3@#@&2mhlLngl:nxrsnNAa|HC	lT+?4kah+	YdZmx^n^Rlk2J@#@&2M.KlT+glh+xEw+[2XmHmxCL?tr2s+UYk/mx^Vcl/aJ@#@&@#@&B&z,b/P&r1@#@&w1\{kO.b1YbWUPx~M+5EdYvJ)^DkWUE*@#@&@#@&vJz~}nA1PGb:)Az?3@#@&mCs^PWanxG4`*@#@&@#@&BJz~?3P,Ku2,o2G2p~}A93/:@#@&/O,W8NsN2XZ^C/kPxPg+A~ams[2XZVmdd@#@&@#@&B&z~!AK~nz/|zM3~&fPHitA3IU@#@&nC13mo+&x6G{&f~',In5!+/DcJhlm0CL+&x6Wmq9E*@#@&?d/bWUKmm3CLqU0Km&f~{PU+/kkKU`rw^b9:rUhlm0Coqx6GmqGJ*@#@&kW~U+d/bGxhl^Vmo+(U6WmqGxrJ~}I,V+	`hCm0lL+&xWG|qf*@*!,YtU@#@&iw1\mkUOhl^3mL+&xWG,'PKC13Co(	0G|qG@#@&i?d/bWU`rw^)9:k	Kl13lTn(x6W|q9Jbxam-{bUYhl^Vmo+(U6W@#@&sk+@#@&dam\|k	OnmmVlT+(U6WP{~?//bGUnmm0lL+(U6WmqG@#@&x[~b0@#@&7@#@&vzJ~w29A(,ZIAfAHK&bJ?@#@&5;DX,xPidJU3J2;K,?4k2hxOKH2+kR;dDq9S,?4kahxO:Xa+/cwmd/SW.NBPj4bw:UY:Xwd b1m/dSr^xd+,E@#@&;;nMXPx~$EnDH~LPEwI}HPUtb2:xOKHwnd,J@#@&5EDX,x~;!+MX~[~EqC3IA~`v`j4bw:nUDKzwdck[Utbw:xDb'8#bird@#@&k+Y,./{/+M-nDcZM+CYn6(LnmDcJzf69~RIn^KD[?Or#@#@&/YPM/{^W	xO+sw nX+m!O+v;E.z#@#@&b0~16P,DdRG0,Y4n	@#@&721\m/D.zm^KE	Y1!:(nD{Dd`rEdnMqfrb@#@&dw1-m/DDt+O+.H!:8+MxDk`E2m//AGMNE#@#@&iw^7{kYDAx7rDKxh+	Yx.k`Jz^m//dr^+	/Jb@#@&n	N~k6@#@&k+O~M/'UGDtrxT@#@&@#@&@#@&Ezz,?AJ2;K~fzK)~U2K@#@&B,@*@*@*~Pl(V/lP2^hl^3mL+&xWG@#@&;;nMX~',7iJjASAZK,w1Kl13CoqUWKRw1Kl13lTn(x6W|q9~~21nCm0CoqUWKRw^KmmVlTn&xWK{:Dl13bUogEh4DS~amnm^3mo+&UWWcw1nCmVCT+(x6G{Utr2a+N9CD+SPr@#@&;;DHP',;!nDHP'Prw^Kmm3mL+&x0K 2mhl13Con(	0G{w9(;l..b+D/G9+~J@#@&$EnMX,'P$E.X,[~JwI6\,wmhCm0lo(U0KPr@#@&;;nMX~',5EDz~LPJ	uAI3Pa^hl^0lT+q	0K w1nCm0lLn&x0KmqG'J,'~w1\|kUYKC13Co(x6W~'rPJ7@#@&B.+k2KxdRSDkD+,5EDz@#@&B.nkwW	d+c+x9@#@&/Y,Dd'dnM\nDc/DlOn}4Ln^D`EbG6GA "+1WD9?OJ*@#@&/Y~.k'mKUxD+:a n6m!Yn`5;Dz#@#@&@#@&kW~grK~.kRnW6~Dtn	di@#@&@#@&E&z,S6rnjK~:C2,Kb;|bV3~qgs}@#@&d2^7{dYMPDmmVr	o1;h(+.'MdvJ21nmm3mo(x6WmKMl^Vbxog;:(+Drb@#@&iw1\m/O.UtrwGCY'.dvJw^KmmVlTn&xWK{Utkaw[fmYnJ*@#@&7am\|dYMZlM.r+MZKNn'.dvJ2mhCm0lLn&x0GmwfpZm.MknMZKN+r#@#@&d@#@&nx9PrW@#@&/OPM/'	GOtbxT@#@&@#@&vJz~?APP"2}i&I29~jb]qz$d2j@#@&am\|/D.HY4W91ChP',EsG(?4r2fVYnIn5!+dYr@#@&am-mkYD\nDtGN"naVz,',JsG(U4kafnVYn]wVHE@#@&ZEkOG:D:DCxdC1YrW	(NxOr6k+.~{P2m7mkY.:Dmm3bxTHEs4nD@#@&rW,wm7m/DDZm..kD;W[+~x,JEP}]Pam-mkYD/CMDr+M/KNn,',1jdS,Otx@#@&iw^-|/YM/lMDk./W9+,'~Jo9oME@#@&nx9PrW@#@&@#@&v?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U=@#@&BP3HG)~nmLPJKl9@#@&EU?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=@#@&@#@&@#@&@#@&veCeMeCeMMCeeCMeCeeCMMeCeCeMeMMCeMeCMeCeMMCeeMMCeMeCMCeMCeCeeCeCMeCeMeCeMMCe@#@&vPUKb"PlPhrUK~A)/n@#@&BCMeCeMMCeeMMCeMeCMCeMCeCeeCeCMeCeMeCeMMCeeCMeCeeCMMeCeCeMeMMCeMeCMeCeMMCeeM@#@&kWPMn$EnkYc0WM:vE/!4hkDJb@!@*JJ,Otx@#@&@#@&dEU?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U@#@&dEP!nDPlss,WWPD4P.;!kDN,rx6W.:mYrG	R@#@&7B?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U7@#@&7@#@&7BJ&PV+UnMkm~nMDGD,WKD~alT+@#@&da^\|/ODV+UnMkmhCo2DMG.P{PrbOPsnm/OPKU+,Dn5!kDn[,0r+^[,hCkP:wDXc~J@#@&7@#@&dv=?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?@#@&dEP;tnmV~6W.PjCVbNCObWx~3MDGDk ,fG,xKYPaDK^+N~k6PO4D+,CDP+M.GDkR@#@&7B==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U@#@&iq0~21\mk	OAD.@*!,Ktx@#@&diDn/aWUdRD[kM+mD~2mhlT+Hlhn,[~JQh/T'E~LPw^-|/ODVn	+.bmhlo2M.WM@#@&dAVdn@#@&di7@#@&ddE==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U@#@&77EP$EbsN,r;.,KDCUkl^YbG	R@#@&diBU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?@#@&diW8%w+NAaZ^l/k H+S(tSPDCUkl^YbGx,w^-|/Y.\Y4W9Hm:nBPam\|/D.b1mGE	YH;s4+MSPam\|dODt+D+.1;h(+.~,2m7{dOMZl..b+.ZK[~~;EkYWs+MPDmxdl1YrG	qNUYb0k.@#@&idiW8Lon92aZ^C/kR	.bY+jr	os+hCM+UDPrKDmm0rxT1;:(+.EBPw1-{kYD:.Cm0k	oHEh8D7d@#@&diW8%w+N3a;VC/k Ax[oHdKDmxkCmDkGx,w^-|/YM\+DtW9HC:@#@&d7@#@&7iB&z,KDbxO~KEY~G!D~xA^X~6WM:+9PMn;!+dY,6hs@#@&divD/wKUd+chMkO+~WNn6|2WkY[CDl@#@&7iB.+k2KxdRxN@#@&i7@#@&d7B?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU@#@&diBPUnUN,r!D~K.C	/CmDrW	R@#@&idB==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?@#@&d71l^VPK4No+92aZ^lddc?+	[(tSI5;+kYv0nNna|wG/D[lDlS~am\mdDD3x7rMWUs+	Y#@#@&i7@#@&d7BJz~KMkxD~W!YPK;.PM+kwGxdn@#@&7dE.+kwGUk+RA.bYnPw3G2p|D/E^Y@#@&diB.+kwGUk+RUN@#@&di@#@&diB?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?@#@&divPdWC[,rE.~"+dwKUk+ @#@&idB?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U@#@&id^C^VPK8Lw+NAa/Vm/kRJWC[oHJIdE^Ydcw2f3p|Dn/!sD#@#@&di@#@&id@#@&diB=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?@#@&i7B,Zt^VP6WMPnD.GM/~0MG:,sn[A6R@#@&idvU?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?didd@#@&i7mmVsPK4%oN2X/Vm//cp\S"+kwGxdnj+.k6z`AD.Kmo+HCs+b@#@&7i@#@&id@#@&diB?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U@#@&77EPI[kM+mD~AkDt,l~HndklL+,6I,mGhaV+On,/G:~Dld0R@#@&diB?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U@#@&77b0Pg6K,V+	c2m7{kY.2..KD\/Tb@*ZPO4x@#@&7i@#@&di7am-|/DDCbNoWM:xJDD;nr@#@&i7dEUU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?@#@&didEPj+O~}E.P"n/aWUdPfCOmPOW,JKmC^R@#@&didE=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?@#@&didv~z\lbsl(V+,\nY4W9/~hrs^Pd+m.m4P;U^k:rON~V-Vd,W6P1KNdP(X~/wC.mYk	LP	WNd~hbY4PCPE&rR@#@&i7dEPqb,I+C["+dwKUk+KmDxY@#@&i7dEP+#,InC9I+k2W	/+gG[+@#@&id7@#@&7idvzJ&zJz&&Jzz&&Jz&zJ&Jz&JzJzzJzJ&zJz&zJz&&JzzJ&zJzzJ&&zJzJz&z&&Jz&zJ&zJz&&Jzz&&Jz&zJ&Jz&JzJzzJzJ&zJz&@#@&d77EP1KO+=PY4nd+,lM+~Y4n,w.ksCDHP-C^E+dS,4;Y,O4+.PmD+,:mUX,:GDP2Gk/k(s+,D+D;.x,\mV;+d@#@&d7dE&zJz&&Jzz&&Jz&zJ&Jz&JzJzzJzJ&zJz&zJz&&JzzJ&zJzzJ&&zJzJz&z&&Jz&zJ&zJz&&Jzz&&Jz&zJ&Jz&JzJzzJzJ&zJz&zJ@#@&7id@#@&7diBzJ~u2zfAI@#@&77iw^\|dYMZ;dDW:n.:DCxkC1YrKx&N+	YbWkD~',W8%w+NAaZ^l/k ]+mN"+dwGUk+HW9n`rz&]wVzul[+MEBPE;EkYWs+MPDmxdl1YrG	qNUYb0k.E#i@#@&d7d@#@&id7BJ&PAI]6"@#@&77iw^\|dDD3MDKDZKN~',W8Lw+[3XZVmd/cI+m[]+kwKxd+HG9+cJJ&2MDG.r~PE/KNnJ*@#@&d7iw1\{kYM3DMW.H/dCT+P{~W(Ls[36;Vm/dR]nmN]+k2W	/nHKN+cEJz3DMGMJS,Jt+/klTnJ*@#@&did@#@&iddOB4JAA==^#~@%>
			
			
			<%#@~^3QgAAA==@#@&d7iBJzP&xknDDP/W9+~O4lY,Ak^VP9ns+D+,Y4+~2mmVlTnPktr2s+xO~bxWW@#@&id7$EDX{JG3SAK3PwI6\,wmhCm0lo(U0KPr@#@&d77$EnDH~',;;nMXP'~ru2"3,w^hl13lT+&U0KR2mhl^Vmo+&U0K{qGxEPLPam-{rUDnCm0CoqUWKP[E~rd@#@&i7i@#@&idi/+DPMd'k+.\D /M+lDnr(L+1OcJzf}f$R]n1W.NUnYr#@#@&idddnDP./{^KxUD+swR6^ED+c;!+.z*@#@&i7dk+Y,.d'	WDtrxL@#@&d7d@#@&didrW,+D. 	Eh4.@!@*T,Y4+x@#@&i7dimCV^P^sK/+98`*@#@&i77dM+kwGxdncDnNb.+1Y~21nlLnglh+,',J_s/T'K4+MnPSldPmx~nMDWM~wMWmddk	o,XGE.~M+5EdYcPKsl/n~DDzPmLmkU,WMPmKxDCmDP/EkYGhDPUnD7km~CY,Fc%T! !KRo+936vIb~R!! *+& &2fORE@#@&iddVkn@#@&d7d@#@&77idBJ&P"+/DG.+,Y4+~b8r^kOX,OW,InRUtk2~Dtr/,2mmVmoR@#@&di7d$EnDH'Eihfb:3PhDW9;^YkrMNnDn[,?3K,2mhD[6MN{j4bw2+9xZ~21nmm3mo(x6WmqG'T~qC2"3Pamnm^VlT+&xWWm(G'EPL~w1\mr	YnC^0lL+&U6W~LPriJ@#@&i7di/nY,Ddxk+D7nDcZDCO+}4N+^YcEzf6f~ ImG.9?+OE*@#@&di7i/nDPM/'1W	UY:2R6n^!Y+v5EDX*@#@&didi/nY~.k'UWD4k	o7@#@&dd77id7d@#@&id7iBJzP;t^3,Y4+,r.[DBk~?4kwarUo,?DlOEd@#@&d7di2m7{dOMrD[nM?OlD;k'f@#@&iddi;!nDH'E?AS3/:PnMGN!mYk6.NDN w^KMN6D9m?4k22NPo]}H~nMG9E^D/}DND[P&1H2"PB6&1P}.ND/,6HPvnMW[E^Okr.N.+9Rr[KDNn.{r.N.kRr9WMN+M#,	CAI3P}D[nM/Rb[WMN+MxEPLPam-{rUDr.N.qGP'~riJ@#@&id7dknDP.k'k+D7+M ZM+CYr8%mYvEbGrf~ ]+1WMNj+OE*@#@&di7dk+O~M/'^G	xO+s2c+am!Y+v;!nDH#@#@&id77b0P	GY,D/cnG0,Y4+Ud@#@&id7di[W,h4r^+PH6:P./cnK0@#@&diddida^\|/OD:+h2}DNU4kaw+9x./vJamKD[6MNm?4rwa+[E*@#@&77id7dbW,w^7{kYD:+s2rMNjtbw2n9'F,Otx@#@&77didid2m-mkY.rM[+M?OCDE/x{@#@&7di7idn	N,k0i@#@&7did7DkRhG7+xaYiddi@#@&dididsWG2@#@&7di7+	N~r6@#@&77idd+D~M/x	WDtk	oi@#@&id7dEz&~`wNmO+,Yt~6D9+MPjYCO!/~YK~Jh+U[bxoE~KD~JhCMYrmV^XPUtb2wNE@#@&d77ik0,2m7{/D.6D9+M?OlO;k'{PD4+	@#@&7idd75!+.X{E`n9zKAPrMN./,?3K,W.[D?DCY!/'F~	CAIAPrNG.9+.'r~[,w^-|kxO6MNnD&9,[~rir@#@&idi7+^/n@#@&d77id;!nDH'J`K9b:2,r.Nn.kPj2:~WMNn.UYlO;k'fPquAI3,k9WD9+MxJ,[~w1\mr	YrM[+Mqf,'~JpJ@#@&7d77x[PbW@#@&d77i/+O~M/x/.7+.cZM+lD+}8LmO`rb96GAR"nmKDNUnOJ*@#@&d7d7dY~DkxmKxUO:w nX+^EDnv;;DH#@#@&di7dk+OPM/xUKYtbUoi@#@&i77d@#@&id7dv&JP/VCD,Y4n,?+ddbWU/@#@&id7iw1/{;VCDzVs?/drKx/vb@#@&ddi7@#@&ididvz&~;VG/~Y4+~/Kxxn^DkGx@#@&id7immVV,m^G/N8`*@#@&7idd@#@&diddE&&P"+9k.+^O,YGPD4+,?4ra:+UO,HCxmLD@#@&diddM+k2W	/nRM+[rM+mD~Jw+NAamHmxmon?4ra:nxDdI/;sD/RCdagrN{E,[~am7{k	Y}.ND(f,[~EL:/Tx5KED,j4ka:xOP4CkP8+UP9+snD+N E@#@&7di7M+daW	/+c+	[di@#@&did77id@#@&7di+x9~r0idi@#@&d77F4UCAA==^#~@%>			
			
			
			<table class="pcCPcontent">
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr>
					<th colspan="2">FedEx<sup>&reg;</sup> Shipment Canceled</th>
				</tr>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>	
			</table>
			
			<%#@~^4QAAAA==@#@&d7x9Pk6@#@&7+	N~k6d~~@#@&d@#@&+	NPbW@#@&EeCeMeMMCeMeCMeCeMMCeeMMCeMeCMCeMCeCeeCeCMeCeMeCeMMCeeCMeCeeCMMeCeCeMeMMCeMe@#@&B,2H9=Pn6j:P$b;F@#@&vCeCeeCeCMeCeMeCeMMCeeCMeCeeCMMeCeCeMeMMCeMeCMeCeMMCeeMMCeMeCMCeMCeCeeCeCMeC@#@&cCQAAA==^#~@%>
<%#@~^QgAAAA==~@#@&hko{D+$EdYc;;+MXdOMkxTcJs/orb@#@&@#@&b0~:dL@!@*EJ,Otx~@#@&dzBAAAA==^#~@%>
	<div class="pcCPmessage">
		<img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"> <%=#@~^AwAAAA==hkoRwEAAA==^#~@%>
	</div>
	<%#@~^EAAAAA==~@#@&n	N,k0,@#@&lAIAAA==^#~@%>
<%#@~^JQAAAA==~b0~am7{/DD_rNsGDsP@!@*,JYM;+rPY4nUPywsAAA==^#~@%>
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">FedEx<sup>&reg;</sup> Delete Shipment Request</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
			<p>
			<b>NOTE:</b>  When shipping with FedEx Ground, you cannot delete a shipment once a Close operation has been performed. 
			<br />
			<br />
			<b>NOTE:</b>  When shipping with FedEx Express, you must delete a shipment prior to the end of business.
			</p>
		</td>
	</tr>
</table>

<table class="pcCPcontent">

	<form name="form1" action="<%=#@~^CgAAAA==21nCT+gl:0QMAAA==^#~@%>" method="post" class="pcForms">
	<input name="PackageInfo_ID" type="hidden" value="<%=#@~^EgAAAA==21\mbxDnl13mL+&xWWKwcAAA==^#~@%>">
	<input name="id" type="hidden" value="<%=#@~^DgAAAA==21\mbxDrD9+M(ffAUAAA==^#~@%>">
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>		
		<tr>
			<th colspan="2">Are you sure you want to delete this shipment?</th>
		</tr>		
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>	
		<tr>
			<td></td>
			<td>
			<input type=submit name="submit" value="Request Delete Shipment" class="ibtnGrey">&nbsp;&nbsp;
			<%#@~^SgAAAA==~@#@&7idam\|/D.nM+-kKEdKmo+,xPrrD9[nYmk^/ ld2Qk['r~[,w^-|kxO6MNnD&9@#@&7idyBUAAA==^#~@%>
			<input type="button" name="Button" value="Go Back To Order Details" onClick="document.location.href='<%=#@~^EwAAAA==21\mkYMnD\bGEknCo2wcAAA==^#~@%>'" class="ibtnGrey">
			</td>
		</tr>
	</form>
</table>
<%#@~^SgAAAA==~x[,k6@#@&EzJ~fA?PI}5~P_2Pw3fA(P}$B2;K@#@&d+O~K4%s[2XZsCk/Px~	WOtbUT@#@&mxMAAA==^#~@%>
<!--#include file="AdminFooter.asp"-->