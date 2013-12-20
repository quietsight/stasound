<%@ LANGUAGE = VBScript.Encode %>
<%#@~^vwEAAA==@#@&BP4kkP0bV~kkP2lMY~G6PnMGN!mY;C.YBPmx~+^Gs:nD1nPmw2sbmlOrKx~N-VGa+9Pl	N,dW^N~4HP3CMVX,(:almD~JS;R,n.W[;1Y/lMO~,kOd,/W;.1+~mK[~~DtPnMW9;mDZCDDPUCs+PmUN,VWTG~lM+,w.W2nMYzPKWPAl.sHPqh2mmO~,JdZ ,ZKwXMkT4Y, T!8O+TZ&R,)V^PDbL4YkPM+d+.-N PIGE,l.n,xWO~mVsWSn9POKP!/+BPmsYDSP9kdOMk4!O+,lx9&GD,D/nVs~mxzPaCDD/~G6Pn.G9E^Y;CMYvkPkWEMm~mKNnPSkO4KEY,OtPhMrOYx,mGxdn	Y~W6~2mDsz,q:2C1Y P:G,mG	YmmY,2m.VHP(:al^OBPw^nlk+P7rdkDPShARnCMVzks2l1Y ^K:R@#@&MpwAAA==^#~@%>
<%#@~^OwAAAA==~alLKbYV'rjtbw2k	o~	bylM[P P?bLUlDEM+~n.GK0~r6~fVr-DXE~sRQAAA==^#~@%>
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
<%#@~^gDwAAA==~@#@&9b:,;EDHSPM/SP1WUUD+:aSPK4Lw392o(sV9W^S,W8LwnNA6jOM+lhS,/ODwr^+Hm:~PVDm2tbmpHd@#@&9b:PbKlT+Z!..+	YBP-l.o^lLq	^WswsnD+~~;Dz~,dDD6"fBPw1\|rxDr.ND(9@#@&fbhPam\|dODt+DtGNHCs+SPa^\|/O.t+Y4G9Inw^zBP/!/DW:D:.l	/CmDkGU&N+	Ok6k+MS~w1\|/OD)^1W;xDHEs4n.BPw^-|/ODtnD+.gEs4+M~,2m7{dYMZC.Mk+M/W9+@#@&9r:,w1\m/O.:DCm0rxT1;h(+DS~am-{kOM?4bws+xDb1^W!xO1!:8nM@#@&Gr:,wm7mdYMf/OkUCDkGx;GE	Y.z;WNnS,w^\|dDD9/DkxmYbGxhWdYmV/G9+~,2m7{/D.Jl	o!lL+/G9+SPa^\|/O.dWmCsZGNS,w^7{kYDG+DCk^?^l	/S~am\|dYMnlTrUo:W0+U@#@&9b:~0[+X{2GkYNCOm~~W(%w+[A6;Vlk/B~W(L6EDw;OoHSGGmBP/M-o2G2o(hVuODwSPw3fA(m./EsOBPo2G3o{i"SBPw1\|dYM2.DKD\dT~Pa^\|/YM)^YbW	@#@&@#@&@#@&@#@&B?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U@#@&EPjKz]K=P6H,Sr)9@#@&vU?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?U@#@&@#@&B&z,!AK~}IG2I,qG@#@&am-{kY.6MN+M(f{I+$;n/D`rk[Jb@#@&w^\|dYM?ndkkWU6MNnD&9{?nk/bWxvJa^b9:rx}D[nMqfrb@#@&k0,2^\|/DDj+ddbWUrM[+Mq9xrJP6],Vnxv21\mkYMrD9+M(f*@*TPDtnU@#@&da^\|kxD6.ND&fxw^-|/OD}.ND(9@#@&djnk/rW	crw^zNskx}D9nD&fE#{w^-|kxD6D9+D&9@#@&Vk+@#@&721\mk	OrMNn.&f'2^7{dYMj/dbW	rD9+M(f@#@&nx9PrW@#@&@#@&vzJPnz!3Pgbt2@#@&2^hlL+gC:'EoN2amtlUlTnUtra:xYk?h6fcldwr@#@&3MDnmL+gl:xEsNA6mHCUmon?4rws+UOk?n69cldwr@#@&@#@&EzJPb;K&61@#@&2m7{dOMbmDrW	P',.n;!+kYcJ)^DkGxrb@#@&@#@&vJzP6KA1~fzPzA)U2@#@&mmV^~Wa+Uf(`b@#@&@#@&E&z,?2:~PCAPw292p~}AB2;P@#@&/nO,W4%oN36;sm/d,',1+SPa^sN36;VCdk@#@&@#@&BJzPV3PPhb;|)M3~&f~1`\AAIj@#@&nl^Vmonq	WK{(GP{PI;!n/D`EnmmVCT+q	WW|qfrb@#@&U+k/rWUKmmVlTnq	0Gm&fPx~U+d/bG	`EamzN:bxhCm0lL+&xWG|qfrb@#@&k0,jn/kkKxKl^Vmonq	WW|q9xrJP6],VnxvKmmVmoqx6W|(f*@*TPDtnU@#@&da^\|kxDKCm0lT+(xWG,'~nm^3mon(	0Wm(G@#@&dUnk/rKxvJw1b9hk	nCm0lLn&x0KmqGJ#{2^\|k	YKl^Vmonq	WW@#@&nsk+@#@&7am-{bUDnC13mo+&x6GP{Pj+k/rG	nl1VlT+q	WG{&f@#@&nx[~b0@#@&i@#@&Ez&~w2f3p,Z]2G3gK(zSU@#@&$E.X,'~diJj3d2Z:~?4kwsnUY:Xa+dR;dD(fB~?4k2hxYPza+dRaCk/AKD9~PUtb2:xOKHwndcbm1n/kSk1nU/Pr@#@&;;nMX~',5EDz~LPJo]}H~?4ra:n	Y:Xw/,E@#@&;;+MX~x,;E.X,[Pr	u2"2,`c`j4bwh+	OKHwndckNj4bwh+	O*'q*#pJd@#@&knY,Dd'k+.-DR;.+mY+}8%+1YvJ)f69~R]+1GD9?nOr#@#@&dY~Dkx1WU	Y:wc+Xnm!Yn`$En.H#@#@&r0,1r:~./c+K0~Y4n	@#@&da^\|/O.zmmG;	YHEs8DxM/vJEk+M(fr#@#@&iw^-|/YM\+D+Dg;h4D{Dd`E2m/dhK.Nr#@#@&iwm-mkY.2	-bDG	:xY{DkcJzm^+k/Jr1+xknJ*@#@&U[Pb0@#@&d+O~M/xxKOtbxL@#@&@#@&@#@&Ez&PU3d2/:PGbKzPU3K@#@&vP@*@*@*~:l4^n/=Pw1KCm0lT+(xWG@#@&5E.X,'~7iJ?3JAZPPa^hl^0lT+q	0K w1nCm0lLn&x0KmqG~Pa^Kl13monqUWKR2mhCm0lLn&x0Gm:DCm0r	oH!:(+DBPa^nmmVlT+(U6WRa^nmm3mLnq	0K{jtr2a+[fmO+BPE@#@&;En.HPxP$;Dz,[,Jw1nm^3monq	0G amnm^3mo+&UWW|sG(/l..b+.ZK[+BP2^hlmVCT+(x6Gcw^hl13lT+&U0K{jtbwPGgl:SPamnm^VlT+&xWW 21nCm0CoqUWK{?4raKGb9[M+dkFBPJ@#@&$;+MX~',;;nMXPL~Jamnm^VlT+&xWW 21nCm0CoqUWK{?4raKGb9[M+dk BPw1nm^3monq	0G amnm^3mo+&UWW|?4k2KG/bYz~,2mhl^Vmo+(U6W w1KmmVmoqx6W|jtbwPWUYCOZW9n~,J@#@&5;+MX,'~;;nMX~[,Ew1nC^0lon(	0GRa^hl^0lT+q	0KmZK:h+	YdS,wmhCm0lo(U0KRamKl^Vmonq	WW|?4rasDGhzYO+	ObWUgls+~,w1Kl13CoqUWKRw1Kl13lTn(x6W|?4k2PKn4W	n~,J@#@&$E+.z,'~;!nMX~LPrwmhl1VlT+(x6W 21nl1VlT+q	WG{UtbwPWtra~~w1Kl13CLqxWGcw^nm^0lLq	0W|?4rw:W/W!xO.H~Pa^nmm3mLnq	0KR2mKC13Co(x6WmoG(ZC.MknD;G9+~r@#@&;EDH~',;;+MX~',Js"6H,wmhC^3moqU0G~r@#@&;!nDHPx~$E+.z,[~JquAI3,w1nl13mL+&xWWcw^Kmm3mL+&x0Km(f{J,[~w^-|kUYhCm0lLn&x0G~LJ~Ji@#@&/nDPM/'k+M-+MR/DlOn}4L^YvJbG69AcImGD[jYE#@#@&/Y~.k'mGU	Yn:a 6n1ED+`$E.X*@#@&@#@&kW~grK,./c+W6~Otxid@#@&v&JPJr}FjhPPuAPI3/&n(2gP,f):b@#@&B,@*@*@*Pk+OPk+ddbWxk~0KDPa.n0bV^+[P[CDl@#@&ij+k/rG	`J2^zNhk	KDdKxgl:J*xDk`Ew1nC^0lo(x6W{U4rw:Wglh+Eb@#@&7?d/bWUcrwm)[skUZKhalUH1m:+r#{EJ,B./vJEb@#@&dUn/kkW	cEw1b9:rx/G	/ro	n+dkUn8J#x.k`Ew1KmmVmoqx6W|jtbwPWzN[.//8E#@#@&dUnd/bW	`Ew^)9:rx;GxkkLU+SrU E#{.k`Eamhlm0lTnq	0G{Utr2:Wb9[D//yEb@#@&dU+d/rG	`Ew1)NskU/Kx/rL	+nZbOHJb{Dk`JamhCm0lL+&xWG|?tb2KKZkDzE#@#@&i?n/drKxcJa^b9:rU;WxdrTxn+UOmYn}DhDW7k	^+;W[+r#x.k`Ja^nmm3mLnq	0K{jtr2:WjYmO+;W[nr#@#@&7U+d/bG	`EamzN:bx;GxkkLx+KGkYl^/W9+J*x./vJamKl^Vmonq	WW|?4raKWtraJb@#@&7U+dkkKx`rw1)NskUZKxdrTx+/W!xYMz/W9+r#xDdcrw^nm^3mon(	0Wmj4k2KK/KEUDDHJ#@#@&ij+k/rW	`E21bNsrxtWD(U0KDslOkGUr#xDkcJamKC13lLn&xWW|/K:hxD/J*@#@&7@#@&B~@*@*@*~dYP^GmmVP7C.km4^+dPWGMP%l7C0bVs~9lYC@#@&d2m7mkY."+1kwb+	O1m:n'M/cEamnm^3mo+&UWW|?4k2KGHm:nJ*@#@&iw^-|/Y.]mrwbn	Yom6gE:(+MxPrJ~BJz.dvJJ*@#@&iwm7mdYM?x[+.Hm:n'Md`rw^Kmm3CLqU0KmUtrasMW:zYDnxDkGxglhnr#@#@&7w1\{kO.?x9+.n4G	+HEs8+M'.dvJw^KmmVlTn&xWK{UtkaKKKtKxnJ*@#@&7@#@&B,@*@*@*P/O~VKmmV~\C.bl8VdP6W.~Nl\CWbVsP9CDl7ididdidi7di@#@&dam-mkYDtCk^I+1r2kxD1C:nxM/cJa^nmmVCT+qUWK{jtb2:WHm:J#idi7did@#@&iw^-|/YM]+1kwdrU+8'M/cJ2^hl^3mL+&xWG|?tr2:W)N9./d8J*@#@&iw1-{kY.Imr2dkx+'M/`r2^nmm0lL+(U6Wm?4rw:W)[9D+ddyJb@#@&7am-|/DDImb2ZbYz'M/cEamnm^3mo+&UWW|?4k2KG/bYzJ*@#@&iw^-|/Y.]mrwUOmYn}DhDW7k	^+;W[+{DdcrwmhCm0lo(U0K{UtrwPGUYCY/W9+Eb@#@&d2^7{dYM]mranK/YmV;GN'./vJ2^hlm0Coqx6Gm?4kaKG}r2r#@#@&i2m7{dOMI+^raZGE	OMX/KN'Dk`r2mhl^3mon(	0W|jtbwKK/GE	YMXE#@#@&i@#@&BJ&Pdr6F`nPPuAPKb;FzM3,qgsr@#@&i2m7{dYMK.C13k	L1!:4.xDk`rw^nC^0lL+&U0K{P.mm3rUT1;:(nMJb@#@&iwm7{kODUtrwGlOn{D/vEw1nl1VCoq	0G{j4bw2+99lD+Eb@#@&d2^7{dYM/mD.b+MZW9+{./vJ2mhl^Vmo+&U0K{sGp/lMDb+.ZG[Jb@#@&nVk+@#@&i@#@&nU9Pr0@#@&k+O,Dk'xKY4rxT@#@&@#@&B&&,?2:~IApj&]3f,.zI(b$JA?@#@&a^\|/O.t+Y4G91C:~{PEwfo?n}f"n;!+dYr@#@&21\{kODt+Y4G[Iw^X~'~Ewfp?h6f"+2sHJ@#@&/!/OWsnMK.mxklmDkKUq9+UYb0rnMP',EnMWN!^OZmDD{jn69rP'Pa^\|/O.:Dl^VbxL1!h(+.@#@&am\|/D.?4k2:xO)1mW!UYgE:(n.'am7{dY.)1mGE	O1!:8nMPB&&,rAx.E/~zm1WE	Y,HEs4nD@#@&rW,wm7m/DDZm..kD;W[+~x,JEP}]Pam-mkYD/CMDr+M/KNn,',1jdS,Otx@#@&iw^-|/YM/lMDk./W9+,'~Jo9oME@#@&nx9PrW@#@&@#@&vJz~szp,I35j&I2GPwJbV@#@&w1\mdDDsmaSYY.r/"+$ErDn[,'~I5E/OcrslaJYO+ME*@#@&EP@*@*@*,DWD/4PDtn~7lV!nPmxXDrh+,Y4+~0G.sPr/,dE(:rO@#@&kW~M+5EdDRWKDs`JkE(hkDJb@!@*JE~Dt+	@#@&i?+kdrW	`rw^b[hbxolXJ+DYn.r#'2^7{dYMom6JYD+Db/"n;!k.+9@#@&n	NPbW@#@&BP@*@*@*PmVSlz/~dY~4m^3,YG~^WmCs@#@&r0,sxcU+k/kKxvEw1b[:bxoCXS+DO+MJ#*@*TPDtx@#@&721\m/D.sm6JnDY+.rkIn;!rM+[,',?+k/bGxvJ2mzNhr	slXJ+DY+MEb@#@&+^/n@#@&7am-{kODwlaJYYn.b/]+$;bDn9P{P0mVkn@#@&+UN,kW@#@&@#@&@#@&B?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==@#@&B,31G)PhCL+,SKl[@#@&v?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U@#@&@#@&@#@&@#@&veCMeCeMMCeeMMCeMeCMCeMCeCeeCeCMeCeMeCeMMCeeCMeCeeCMMeCeCeMeMMCeMeCMeCeMMCee@#@&EPjKz]:)~hrUKP~b;F@#@&BMeCeMMCeeCMeCeeCMMeCeCeMeMMCeMeCMeCeMMCeeMMCeMeCMCeMCeCeeCeCMeCeMeCeMMCeeCM@#@&k0,.n;!+kY 0G.s`E/!8:bYEb@!@*JE~Dtnx@#@&@#@&7EU?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U@#@&iB~!YPmsV,W0,O4+,D;;k.n9Prx6GDslOrKxR@#@&iB=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?Ud@#@&i@#@&iB&z,MnUDk1~+MDWM~WWMPalL+@#@&iw^\|dYMMnUDk^Kmon2M.KD~{PrbY,VC/DPGxP.n$EkMnN,0ks[PSlkPn:2OHR~J@#@&db0~21\{dOMsC6dnDYnMkkI+$Eb.+9PxPrY.;JPD4+	@#@&i72m7{kY.MnUDrmhCo2..KDPx~am-{kOMMn	+MkmhlTn2MDGD,[~E@!4D,&@*IWE,h;/DP6ksV~C^V~D5EbDn[,0kns9/~sma,SnDYDP6ksNkPAtx~O4+PUKrGPsma~SYD+.PG2DkGx,r/,m4n13+[ r@#@&dU9Pr6@#@&d@#@&dE&z,?OlMY~jD\.OUkN~-l^k9lOkGUkd@#@&i2mk{#C^kNCOKn6Dob+s9drn+M/KU1m:nJBPO.!+~,f*@#@&da^d{jl^k[lOn:+aYwr+^N7E;W:2C	XHlsnr~~6l^/+BP2X@#@&@#@&damdmjlVb[lD+KaOsb+^N7J/G	/ro	n+dkUn8J~~OMEn~,fl@#@&iw1/{jl^rNmYnK6Oob+V97J;WxkrLx+dkU++EBPWl^d+BPfX@#@&d2^k{#l^r9lOK6YwksNiJ/W	/rL	++;rYHJ~,O.E~,!~B&&,s9(AR&l~~oG(MR+Z@#@&da^k{#mVbNlD+:n6Dsr+^N7E;Wxkro	++UOCYrMn.W-r	mnZK[+r~~OME+S~y@#@&da^k{#mVbNlD+:n6Dsr+^N7E;Wxkro	++hGdYmV;W[+ES,Y.ESP8v@#@&iwmdmjlsk9CD+P6DskV97J;WU/boUnZW!UYMXZK[nJBPDD;+S~y@#@&da^/|.CsbNlOn:+aYwrV[iJtWDq	WWM:CYbWUEBP0ms/~P2T@#@&i@#@&d2mdmjlsk9CYKnaDskns9dEI^bwrxD1ls+rSPam-{kY.om6SOYDkk]n;!kM+[~~fl@#@&da^/|.CsbNlOn:+aYwrV[iJ"+mbwbnxDsC6gEh8DJB~w1\{kO.sm6d+OYn.b/]+$;kM+[S,Fv@#@&iw^/|#mVr9lD+K6DokV[dr?nU9+DgC:J~,2^\|/DDolaJYO+Mr/"+5;bD+[S,&X@#@&7amd|.mVk9lDnK6Osb+s[iJ?UNDn4GU+gEs4nDES,w^\|dYMsCad+YOnMkdI5!k.NBPF+@#@&@#@&iw^/|.CsbNlDnK6YwrnV9drHCks]mrwbnxD1ChJ~~21\m/D.wlad+DY+Mkk]+$ErDNS~2*@#@&7w1/{jCsk9lD+P+aOwknV97J"+^raSkUn8JSPa^7{dDDwl6d+DO+MkdI;;rM+NB~&l@#@&i2^/|.mVrNCOKn6DokV[7rI+^raSrx+r~~6l^/+BP2X@#@&d2mk{#C^kNmO+:+6Dor+^NiJ]+^raZrYHE~,w^-|/Y.om6J+DODrkI;EbD[~,!~BJz~oG(2 f*BPsGp!Oy!@#@&7w^d|.CVb[lD+PnXYsrn^N7J"n1k2UYmY+}Dh.W7kUmZG[J~,2m7{/D.olXSYO+.rkIn;!rDNS~y@#@&721/m.msbNCD+:+6DsbnV9dEImr2hW/DCV;WNESPam7{dY.om6J+DO+Mkd];Er.NSP8@#@&7amk{.mVb[lD+P+XYorVNiEImka/GE	YMX/W[nr~~w1-{kY.om6SnOD+.kk];;bDN~, @#@&d@#@&7B?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?@#@&iv~Z4+13~0G.,.CVb[lDkGU,2D.GM/ PGG,xGDPaDW1+[Pb0~Y4+.n,lD~+MDWMd @#@&dEU=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=@#@&d(W,wm-mbxO2M.@*!~:tx@#@&di.+kwGxk+ .NkMnmDPw1KCo1m:nP'~rgh/TxJ,[~21\{dOMMnx.bmKmo2DMWM@#@&i2s/@#@&7id@#@&7dEUU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?@#@&d7B~$!ksN,6EMPP.mx/C^DkGxc@#@&d7EU?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?@#@&idG4Nsn[A6Z^C/kR1ApHdKMlU/C^DkGx,2m7{dOMH+O4KNHlsnBP21\|/YMb1^W!xO1!:8nM~Pa^\|/YM\nYDgEh4n.BP2m7m/DD/CMDkn.;W[+B~;EdDWs+D:DmU/mmOkKx([xYbWkD@#@&77d@#@&id7W8%w+[2X/Vm/d qDkOnUkUo^nhl.xDPJUtb2fmYnJBPG8Ns+936;Vlkd w10|snN3aGlO+wGDslOcam\mdDDjtb2GlO#@#@&didK8Lw+[2XZsCk/Rq.kD+?bULVnmDnxO~rK.l1Vk	oH;s4+.EBP2m7mkY.:Dmm3bxTHEs4nD@#@&77iW4No+926;sC/kRqDrYnjbxLVKlM+UO,JSCUTECo/KNnr~,J2gJi7did7@#@&d77K4LwnNA6Z^Cd/cMkO+jr	os+hCDxO~r?tr2s+UYz^1W;	YgE:(+ME~,w^\|/O.Utkah+	Yb1^GE	YgEh4n.@#@&7diG4Nsn[A6ZsCk/ MrD+jbxTV+hlMnxDPEIY;.	S+DO+MsWMhCYr~,JKfoE@#@&7di@#@&id7G(Lsn[A6/VmdkR	MkD+nmDUY,J/W	0rL!Dl(s+VDW!U[ZKxkkLxnnr~~Jr@#@&id77K4Lon92aZ^Ck/ zN91+S1K[+,JK+M/GUgl:E~,?+kdrW	`rw^b[hbxK+MdW	1ChJ#@#@&id7dK8Nsn92XZVm/k b9NH+S1G[PJ;G:alxHHC:JBPj+ddbWU`r2mzNhr	ZWh2mxz1mhJb@#@&iddidK8Lw+[2XZsCk/Rq.kD+nm.nxDPrb[N.nk/E~,EJ@#@&77idd7G(Lo+93XZsm/kRb9NgnhgW[+,JJr	+FrSPU+/krGxvJam)Nhr	ZGxkro	+nJbx+qE*@#@&di7id7K4Ns+92X/Vm/dRzN[Hh1K[+,JSbUn r~,?n/drKxcJa^b9:rU;WxdrTxn+dr	++r#@#@&didi7dK4%sN3a;VlkdRzNNgnA1KNPEZrOHJSPUn/kkGUvJw^)9:rx;G	/rTx+ZbYHE#@#@&7did77K4LwnNA6Z^Cd/cb9NH+AHKNnPrjYmYn6MnDG-bx^+;G9+EBPU+/kkKU`rw^b9:rU;Wxkro	++UOCYrMn.W-r	mnZK[+r#@#@&idd77iW8Lwn92a;Vm//cb9[1hHW9+~EhW/DCV;WNESPU+k/rWUcrw^b9hk	ZGUkkoUnnG/DC^ZG9+r#@#@&di7didG4Nsn[A6Z^C/kRb9[H+S1KNnPE/KEUYMzZKNnEBP?ndkkGxvEam)9:bxZKxkro	+nZKEUOMXZK[+r#@#@&77didK4%sn[A6/Vmd/c.rD+nC.xOPr)9N./kJ~,JJE@#@&d7diW8%w+NAaZ^l/k )N91hHW[n,J\WMnq	0G.slYrG	JSPUnk/rKxvJw1b9hk	HGDqUWKD:mOkKxJ*@#@&didK4%sn[A6/Vmd/c.rD+nC.xOPr/KxWbo!Dl(V!DKEUN;WUdboxnJBPJJE@#@&idi@#@&d77b0~w1-{kY.om6SnOD+.kk];;bDNP{PD.EPOtx7@#@&ddi7didd@#@&7didK4%sn[A6/Vmd/c.rD+nC.xOProm6JYD+Dr~,EJ@#@&7did7G(Ls[2XZVmddRzN91nhHG9+~J"nmbwrn	Y1ChJSPUnk/rKxvJw1b9hk	Inmbwrn	Y1mh+r#@#@&77didK4%sn[A6/Vmd/cb[[g+hHG9+~J"n1k2b+	Ysm6g;:(+.JBPjnk/kKU`rwmz[hk	Imrwrn	YolXHEs4n.r#@#@&7id7dK8Nsn92XZVm/k b9NH+S1G[PJUnx9+DgCh+r~,?n/drKxcJa^b9:rUU+x[nM1C:E*@#@&ididdK4No+92aZ^lddcbN9H+S1W9n~JU+	NnDK4Kxn1!h4DES,?+ddbWU`r21b[sk	?+	N.n4WU+gEh8DJ*7diddi77didid@#@&77id7dK8Lw+[3XZVCdkR	DbOnCM+	YPrHmrV"+^kaknUDJ~,EJ@#@&di77didK4%sn[A6/Vmd/cb[[g+hHG9+~J"n1k2b+	Y1m:E~,?n/kkGUvJw1)NskxtCrV"+1k2knUD1C:E#id77idd7@#@&d7di7id7K4Ns+92X/Vm/dRqDrOnlMnxDPJz[[D/kJSPEE@#@&7di7did77K4Lon92aZ^Ck/ zN91+S1K[+,JJk	+qEBP?d/bWxvE2mzNskUIn^bwJk	nFr#@#@&idd77id7dK8Nsn92XZVm/k b9NH+S1G[PJdrx JB~j+k/bWU`E21b[:bUImr2dkxn+r#@#@&i7id7ididW(LwnNA6/Vm/d zNNgnhgWN~EZbYHJSPjnk/rW	cJam)[skx]n1k2ZbOHJb@#@&iddidi7diW8Lw+[3XZVmd/cbN9HnhgW9+~JjOmYnrMKDK\rU1+ZG[JSPUnk/rKxvJw1b9hk	InmbwjOmY+}.nMW\bU^+;W9+E#@#@&id7di7didG8Ns+[3XZslkdcb[91h1KN~JhWdYmV/G9+JB~?//bGU`rw1b[:rU"+^kaKWkYCs;WNnE*@#@&di7id7idiW4Ns[2XZslk/ )9N1A1KN+,E/W!xDDzZG[JSPUn/kkGUvJw^)9:rx"n1k2;W!xYMX;GNJb@#@&d77iddi7W(Ls[36;Vm/dR	.bYnnm.+	Y~EzNN.nk/E~,EJJ7ididdid@#@&did7diW8%w+NAaZ^l/k 	DbYnCDnUDPEHmrV"+^rak+UOr~~JJEid7ididdidi7di@#@&did7G(Ls[2XZVmddRqDbYnnC.xOProlXSnOD+DES,J&Ji@#@&d7id@#@&didUN,kWdid77,@#@&i7di@#@&i7G4NsN36/sm/dRAUNoHJPMlxdC1YrW	~am-|/DDHY4GNglh+@#@&77@#@&divzJPnMrUY,W!Y~W;.,xnh^zP6W.hNP.n$En/D~X:s@#@&idBM+k2W	/nRSDrOP0[+X{wKdONmYm@#@&d7vM+dwKU/RnU9@#@&77@#@&7dE=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?@#@&idE~?x[P}E.~:Dl	dl1YkKU @#@&diB=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=@#@&7d1lss,W4%oN36;sm/dc?xNoHd]+$En/D`Wn9+6|2WkYNmOC~,w1\m/O.Ax-kMGxs+UO*@#@&77@#@&7dE&JPKMk	YPKED~W!D~D/2G	/+@#@&diBDd2W	/RADrOPo2G3(|Dnd!VY@#@&idvDdaWUk+c+x9@#@&7d@#@&7dEU==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?@#@&idB,JGl9P}E.P]nkwGxknR@#@&77EUU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=@#@&7imC^V,W4Ns[2XZslk/ JKlNo\S"+/!sO/vsAf3(m./;VDb@#@&d7@#@&dd@#@&idvU?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?@#@&ddEP;4+13~0KD~nMDWMdP6DWs~o+92XR@#@&77EU=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?d77i@#@&771lsV,G(LoNA6Z^lkdRoHJI/2G	/+jnDb0Xv3.DhlT+Hlhn*@#@&di@#@&id@#@&idB==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?@#@&d7EP"+NbD^Y,hrY4PC~t+/kCoPr"~^Wsw^+O+~dK:nPDC/0R@#@&idB==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?@#@&d7b0,1r:P^nxvw^\|/O.ADDK.Hko#@*T~Y4+	@#@&d7@#@&d7@#@&7diB==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?@#@&77iB~?O,r;MP"+/aW	d+,fCYmPOG,SW1CVc@#@&i77B?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?@#@&d77EP)\mrVm4sn,H+O4KNdPSr^V~k+mDm4P!UVb:rYN~s\+^dPK0PgG[+kP(X~/n2mDCYbUo,xG[/PArDt~l,EJJ @#@&iddEP8bP"+CN"+d2Kx/KlM+xD@#@&didEP+#~]l[IdwKxdngWNn@#@&d7d@#@&id7EzJzzJzJ&zJz&zJz&&JzzJ&zJzzJ&&zJzJz&z&&Jz&zJ&zJz&&Jzz&&Jz&zJ&Jz&JzJzzJzJ&zJz&zJz&&JzzJ&zJzzJ&&zJzJz&@#@&7idvPgGY)~O4+/n~mDnPD4P2MkslDHP7CV!+d~,4;O,Yt.+,lD~hl	X,:GDn~aWd/b8VP.nDEDU~7lsEd@#@&7idEzzJzJ&zJz&zJz&&JzzJ&zJzzJ&&zJzJz&z&&Jz&zJ&zJz&&Jzz&&Jz&zJ&Jz&JzJzzJzJ&zJz&zJz&&JzzJ&zJzzJ&&zJzJz&z&@#@&d7d@#@&didv&JPC3)G2]@#@&7id21\|/YMZ!dYK:nD:DCUklmDrW	qNUOk6kD~'~G(Lo+936;VCdkRInC9In/aG	/ngW9+`rzJ]+aVzCl[nMJ~,EZ!/YKhnD:Dmxdl^ObWUq9nxDkWrDJb7@#@&7di@#@&d7iBJzPAI"6I@#@&7diw^-|/YM3DMWD;G[+,',W8Lon92aZ^C/kR]nmNIndaWU/HKNnvJJz2MDK.JBPEZKNnE*@#@&i7dam\|dODADMW.HndklL+,xPK4%oN2a/^ld/c]l["+kwW	/HW9+cJJz3.MWDrSPrH+kdCoJ*@#@&d77@#@&7divzJP/6tsq]\zK(rg~gj\~2"@#@&idi2m7{dYMsCa;Wx6rDslYbGU1!:(+.Px~K4%s[2XZsCk/R]nmN]+k2KxdnmD+	Yv2m7{dYMHnO4WN"nw^X~,EolXZKxWk.hmYrW	HEs4n.r#@#@&7id2m7mkY.d+DY+MP{~W(Lo+92a/^l/k IlN"ndwKxk+Kl.n	Ycw1-{kY.\YtG["+2VHS,JJYD+Dr#@#@&did@#@&id7r6Pw1-{kYDdnOYD@!@*EJ~O4+U@#@&7didvx{''xx{'x'{x{'x{'{''{'{@#@&id7dEz&~UYlMOPdl4s~fmKNrxL@#@&d7div'{'xx{''xx{'x'{x{'x{'{'@#@&di7dEz&P;DnCD+Po\S,0WM~Jl(+^P@#@&77idvKMCm0kULgE:8nM~~2	^KNn9Sm4+^?D.k	oSPwksn:XwSPwkVK.+wkX@#@&d77iW8LwnNA6/sm// HhpHdJm4n^PU+/kkKU`rw^b9:rU:Dl1Vk	o1!h8+MJ*~2m-mkY.SOYDS~rnfoEBPE?h6GJ@#@&di@#@&idi7BJz~SKl[~^l4sP6DWs~OtPM+5EndDPdYMnls@#@&7idd^C^V~W(%w+[A6;Vlk/cJWmNpHdSC8V`V.latk1p\S*@#@&d7@#@&7id7BJ&P`/n~zfr~dDDnls~DW~kl7+PDt~4bxCDHP[CDl@#@&7didW(%o+92XZslddc?C\$k	l.zdl4ns@#@&7di7E'x{'{''{'{x'{'x'{'xx{''@#@&diddE&&PAx9PJl8n^P9+1GNbxL@#@&dd77E'x'{x{'x{'{''{'{x'{'x'{@#@&7iddi@#@&iddiv&z,?m\nPO4Pjn}9PwVCL@#@&d77i;;+Mz{JihfzK2,w1Kl13CoqUWKP?APPamnm^VlT+&xWWmoG(jn}9s^lLxEFB~E@#@&7di7$EnMX,'P$E.X,[~JqC3]APw1Kl13lTn(x6Wcw^nC^0lL+&U0K{(9{JP'~U+d/bG	`EamzN:bxhCm0lL+&xWG|qfrbPLJPr7@#@&ididd+O~M/x/.\D /M+lOn}4%+1OvJ)GrGAR"+1GD9?nYr#@#@&iddid+DPDkx^W	xD+hw nX+^EDn`$En.H#@#@&7id7/O,Dd{xKYtbxT@#@&id7d1lss,mVKd+94`*@#@&didikWPn.MRUEs8+M@!@*T,YtnU@#@&7di7iDnkwKx/RMnNbDnmDP2^hloHls+PL~Egs/T'Ptn.PAlk~l	Pn.MWD~dm\rxT~HW;MPUnrGR,KVld+,Y.z,lomrx,VlDn.Rr@#@&d7d7n^/n@#@&7did721/{/sl.b^sU+dkkKx/v#@#@&did7dM+d2Kx/ DNkMn^Y,Jw+[2amtlUlTn?4k2hxYd]/;VDdcldagbN'rPL~w1\mk	Y6.9+D&9PLPJLhdo{5KE.PjK}f~kk~xKh~C7lksC(VnP7rmPO4+,@!E@*.bnh,?KrG@!&;@*PVbU3cJ@#@&77didM+dwGUk+ +	[did77@#@&d77i+UN,r6d@#@&didd@#@&i7dx[Pb0@#@&iddSWoTAA==^#~@%>
			<table class="pcCPcontent">
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr>
					<th colspan="2">FedEx<sup>&reg;</sup> Signature Proof of Delivery Successfully Requested</th>
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
<table class="pcCPcontent">
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<th colspan="2">FedEx<sup>&reg;</sup> Signature Proof of Delivery Request (SPOD) </th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
			<p>
			The Signature Proof of Delivery (SPOD) offers you the ability to request proof of shipment delivery and the signature of the party which accepted the package in the form of a Letter or Fax Letter.			</p>		</td>
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
			<th colspan="2">SPOD Recipient Display Information</th>
		</tr>		
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td>Ground Consignee Person Name:</td>
			<td>
			<input name="PersonName" type="text" id="PersonName" value="<%=#@~^JQAAAA==210mwk^VsKDsokV[`rnn.kWxgC:J~,O.E#MA0AAA==^#~@%>">
			<%#@~^JwAAAA==21/m"+$EkM+9(:monKmo~Eh+DkGxgl:ESPDD!+TQ4AAA==^#~@%>
			</td>
		</tr>	
		<tr>
			<td>Ground Consignee Company Name:</td>
			<td>
			<input name="CompanyName" type="text" id="CompanyName" value="<%=#@~^JwAAAA==210mwk^VsKDsokV[`rZGhalxHHls+JB~Wl^/#2w0AAA==^#~@%>">
			<%#@~^KQAAAA==21/m"+$EkM+9(:monKmo~E;W:aCxH1lsnE~,0mVd++A4AAA==^#~@%>
			</td>
		</tr>			
		<tr>
			<td>Line 1:</td>
			<td>
			<input name="ConsigneeLine1" type="text" id="ConsigneeLine1" value="<%=#@~^KQAAAA==210mwk^VsKDsokV[`rZGUkko	n+dkxqE~,YMEn#jA4AAA==^#~@%>">
			<%#@~^KwAAAA==21/m"+$EkM+9(:monKmo~E;Wxkro	++drU+8JBPOD;nqQ8AAA==^#~@%>
			</td>
		</tr>	
		<tr>
			<td>Line 2:</td>
			<td>
			<input name="ConsigneeLine2" type="text" id="ConsigneeLine2" value="<%=#@~^KgAAAA==210mwk^VsKDsokV[`rZGUkko	n+dkx+E~,0mVd+b2A4AAA==^#~@%>">
			<%#@~^LAAAAA==21/m"+$EkM+9(:monKmo~E;Wxkro	++drU+yJBPWlsd9Q8AAA==^#~@%>
			</td>
		</tr>	
		<tr>
			<td>City:</td>
			<td>
			<input name="ConsigneeCity" type="text" id="ConsigneeCity" value="<%=#@~^KAAAAA==210mwk^VsKDsokV[`rZGUkko	n+;kYHESPDD!+bbA4AAA==^#~@%>">
			<%#@~^KgAAAA==21/m"+$EkM+9(:monKmo~E;Wxkro	++;rOXr~,Y.EniQ8AAA==^#~@%>
			</td>
		</tr>	
		<tr>
			<td>State Or Province Code:</td>
			<td>
			<input name="ConsigneeStateOrProvinceCode" type="text" id="ConsigneeStateOrProvinceCode" value="<%=#@~^NwAAAA==210mwk^VsKDsokV[`rZGUkko	n+UYlDn6DhDK\rx^n;W[+rSPDD;n*VhQAAA==^#~@%>">
			<%#@~^OQAAAA==21/m"+$EkM+9(:monKmo~E;Wxkro	++UOCYrMn.W-r	mnZK[+r~~OME+cxUAAA==^#~@%>
			</td>
		</tr>	
		<tr>
			<td>PostalCode:</td>
			<td>
			<input name="ConsigneePostalCode" type="text" id="ConsigneePostalCode" value="<%=#@~^LgAAAA==210mwk^VsKDsokV[`rZGUkko	n+hW/DCsZKNJSPO.!+bwRAAAA==^#~@%>">
			<%#@~^MAAAAA==21/m"+$EkM+9(:monKmo~E;Wxkro	++hGdYmV;W[+ES,Y.E3hEAAA==^#~@%>
			</td>
		</tr>	
		<tr>
			<td>CountryCode:</td>
			<td>
			<input name="ConsigneeCountryCode" type="text" id="ConsigneeCountryCode" value="<%=#@~^LwAAAA==210mwk^VsKDsokV[`rZGUkko	n+;WE	O.X;W9+E~~OMEn#QhEAAA==^#~@%>">
			<%#@~^MQAAAA==21/m"+$EkM+9(:monKmo~E;Wxkro	++;G;xDDHZGNnEBPOD!nXxIAAA==^#~@%>
			</td>
		</tr>	
		<tr>
			<td>More Information:</td>
			<td>
			  <textarea name="MoreInformation" cols="40" rows="4" id="MoreInformation"><%=#@~^KwAAAA==210mwk^VsKDsokV[`rHG.qx6GDslYbGUJBP6ls/nbnA8AAA==^#~@%></textarea>
			<%#@~^LQAAAA==21/m"+$EkM+9(:monKmo~EtWD(x6WDsCOkKxr~~0Csk+uRAAAA==^#~@%>
			</td>
		</tr>									
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>	
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer">
			<SCRIPT LANGUAGE="JavaScript" TYPE="text/JavaScript">
			<!--
			function FaxSelected(){
			
			var selectValDom = document.forms['form1'];
			if (selectValDom.FaxLetter.checked == true) {
			document.getElementById('FaxTable').style.display='';
			}else{
			document.getElementById('FaxTable').style.display='none';
			}
			}
			 //-->
			</SCRIPT>
			<input onClick="FaxSelected();" name="FaxLetter" id="FaxLetter" type="checkbox" class="clearBorder" value=true <%#@~^QwAAAA==r6P21\|/YMsmaSYO+Mkd];Eb.+9'JD.;+rPDtnx~./2W	d+ch.rD+`E^4+^3[r#8BgAAA==^#~@%>> 
			I would like a <b>SPOD Fax Letter</b> sent to the following Recipient.</td>
		</tr>
		<tr>
			<td colspan="2">
				<%#@~^wwAAAA==@#@&d7idb0Pam7m/DDolXSnOD+DbdI;Eb.nN@!@*rY.EnE,Y4+	@#@&id77iwm-mkY.fbdaVCH?DXV'rdYHVn'rJ[rkwVmz)	WxEEJ@#@&id7dnsk+@#@&i7did2^7{/O.Gkdw^CH?OHV'JkYHs+{JENb/2smX)7r/b4VEEJ@#@&id7dnU9Pr0@#@&did76DYAAA==^#~@%>
				<table class="pcCPcontent" ID="FaxTable" <%=#@~^EwAAAA==21\mkYMfkkw^CXUYzV6AcAAA==^#~@%>>
					<tr>
						<td colspan="2" valign="top">
							<p class="pcCPnotes">
								<b>NOTE:</b>  When the <b>SPOD Fax Letter</b> option is checked, SPOD Fax Letter form fields are required.
								<br />
								<a href="JavaScript:generateForm();">Click Here</a> to fill the required fields with your default "Ship To" information.
										<script language="JavaScript">
										function generateForm(){
											document.form1.RecipientName.value = "<%=#@~^FAAAAA==21\mkYMI+1kar+	YHls+JQgAAA==^#~@%>";
											document.form1.RecipientFaxNumber.value = "<%=#@~^GQAAAA==21\mkYMI+1kar+	YolX1;h(+DLAoAAA==^#~@%>";
											document.form1.SenderName.value = "<%=#@~^EQAAAA==21\mkYM?+	N.1m:n4wYAAA==^#~@%>";
											document.form1.SenderPhoneNumber.value = "<%=#@~^GAAAAA==21\mkYM?+	N.n4WU+gEh8DxQkAAA==^#~@%>";
											document.form1.MailRecipientName.value = "<%=#@~^GAAAAA==21\mkYMHlbV"nmbwr+	YHCs+qAkAAA==^#~@%>";
											document.form1.RecipLine1.value = "<%=#@~^EQAAAA==21\mkYMI+1kaJk	+qrQYAAA==^#~@%>";
											document.form1.RecipLine2.value = "<%=#@~^EQAAAA==21\mkYMI+1kaJk	++rgYAAA==^#~@%>";
											document.form1.RecipCity.value = "<%=#@~^EAAAAA==21\mkYMI+1ka/kDXjQYAAA==^#~@%>";
											document.form1.RecipStateOrProvinceCode.value = "<%=#@~^HwAAAA==21\mkYMI+1kajYmYnrMn.G7kx1nZKN+dwwAAA==^#~@%>";
											document.form1.RecipPostalCode.value = "<%=#@~^FgAAAA==21\mkYMI+1kaKWkYCV;W[n4ggAAA==^#~@%>";
											document.form1.RecipCountryCode.value = "<%=#@~^FwAAAA==21\mkYMI+1ka/W!xODHZG[YwkAAA==^#~@%>";
										}
										</script>
						</td>
					</tr>			
					<tr>
						<th colspan="2">Fax Letter Recipient</th>
					</tr>	
			
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>		
					<tr>
						<td>Recipient Name:</td>
						<td>
						<input name="RecipientName" type="text" id="RecipientName" value="<%=#@~^PgAAAA==210mwk^VsKDsokV[`rIn^bwkUYgl:ESPam7{dY.om6J+DO+Mkd];Er.NbSRcAAA==^#~@%>">
						<%#@~^QAAAAA==21/m"+$EkM+9(:monKmo~E"+mb2kxYgCh+r~,w^\mdDDolXJ+DYn.b/In5!k.+9ZhgAAA==^#~@%>
						</td>
					</tr>		
					<tr>
						<td>Recipient Fax Number:</td>
						<td>
						<input name="RecipientFaxNumber" type="text" id="RecipientFaxNumber" value="<%=#@~^QwAAAA==210mwk^VsKDsokV[`rIn^bwkUYwl6g;h4Dr~~w^-|/ODwC6d+OODkd];;kMn9#UBkAAA==^#~@%>">
						<%#@~^RQAAAA==21/m"+$EkM+9(:monKmo~E"+mb2kxYwCa1!:(+.JS~am-{kODwlaJYYn.b/]+$;bDn9bRoAAA==^#~@%>
						</td>
					</tr>		
					<tr>
						<td>Sender Name:</td>
						<td>
						<input name="SenderName" type="text" id="SenderName" value="<%=#@~^OwAAAA==210mwk^VsKDsokV[`r?nU9+DgC:J~,2^\|/DDolaJYO+Mr/"+5;bD+[bBxYAAA==^#~@%>">
						<%#@~^PQAAAA==21/m"+$EkM+9(:monKmo~EU+x9nDgl:ESPam7{dY.om6J+DO+Mkd];Er.NJBcAAA==^#~@%>
						</td>
					</tr>		
					<tr>
						<td>Sender Phone Number:</td>
						<td>
						<input name="SenderPhoneNumber" type="text" id="SenderPhoneNumber" value="<%=#@~^QgAAAA==210mwk^VsKDsokV[`r?nU9+Dh4W	+1!h8+MJBP2m-mkY.smaSYOnMk/]n$ErD[*6RgAAA==^#~@%>">
						<%#@~^RAAAAA==21/m"+$EkM+9(:monKmo~EU+x9nDhtW	nHEs4DE~~21\m/D.sm6JnDY+.rkIn;!rM+[BhoAAA==^#~@%>
						</td>
					</tr>
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>		
					<tr>
						<th colspan="2">Fax Letter Recipient Address</th>
					</tr>		
					<tr>
						<td colspan="2" class="pcCPspacer"></td>
					</tr>	
					<tr>
						<td>Recipient Name:</td>
						<td>
						<input name="MailRecipientName" type="text" id="MailRecipientName" value="<%=#@~^QgAAAA==210mwk^VsKDsokV[`rHCr^I+1rwb+xDHC:JBP2m-mkY.smaSYOnMk/]n$ErD[*zBgAAA==^#~@%>">
						<%#@~^RAAAAA==21/m"+$EkM+9(:monKmo~Etlk^]+1kwbnUYgls+E~~21\m/D.sm6JnDY+.rkIn;!rM+[6RkAAA==^#~@%>
						</td>
					</tr>		
					<tr>
						<td>Recip Line 1:</td>
						<td>
						<input name="RecipLine1" type="text" id="RecipLine1" value="<%=#@~^OwAAAA==210mwk^VsKDsokV[`rIn^bwSbU+8J~,2^\|/DDolaJYO+Mr/"+5;bD+[b0RUAAA==^#~@%>">
						<%#@~^PQAAAA==21/m"+$EkM+9(:monKmo~E"+mb2Sbx+8ESPam7{dY.om6J+DO+Mkd];Er.N7hYAAA==^#~@%>
						</td>
					</tr>	
					<tr>
						<td>Recip Line 2:</td>
						<td>
						<input name="RecipLine2" type="text" id="RecipLine2" value="<%=#@~^JgAAAA==210mwk^VsKDsokV[`rIn^bwSbU+yJ~,WCVk+*MA0AAA==^#~@%>">
						<%#@~^KAAAAA==21/m"+$EkM+9(:monKmo~E"+mb2Sbx+yESP6l^/nTQ4AAA==^#~@%>
						</td>
					</tr>	
					<tr>
						<td>Recip City:</td>
						<td>
						<input name="RecipCity" type="text" id="RecipCity" value="<%=#@~^OgAAAA==210mwk^VsKDsokV[`rIn^bwZbOXr~Pa^-{kYMsC6JnDYnDbdI;;rM+NbsRUAAA==^#~@%>">
						<%#@~^PAAAAA==21/m"+$EkM+9(:monKmo~E"+mb2ZbYXrS~w1\|/ODoCXSnYDnDb/]n$Ek.n9zhYAAA==^#~@%>
						</td>
					</tr>	
					<tr>
						<td>State Or Province Code:</td>
						<td>
						<input name="RecipStateOrProvinceCode" type="text" id="RecipStateOrProvinceCode" value="<%=#@~^SQAAAA==210mwk^VsKDsokV[`rIn^bw?DCYrDh.G\bx1+/W[nr~~w1-{kY.om6SnOD+.kk];;bDN#mxsAAA==^#~@%>">
						<%#@~^SwAAAA==21/m"+$EkM+9(:monKmo~E"+mb2?DlY6.nMW7kUmn/KNnJB~w1\mdDDsCad+OY.b/];!kDNuBwAAA==^#~@%>
						</td>
					</tr>	
					<tr>
						<td>Recip PostalCode:</td>
						<td>
						<input name="RecipPostalCode" type="text" id="RecipPostalCode" value="<%=#@~^QAAAAA==210mwk^VsKDsokV[`rIn^bwnKdYmVZK[nJBPam-{dOMsC6dnYD+.rkI+5;bDnN*BhgAAA==^#~@%>">
						<%#@~^QgAAAA==21/m"+$EkM+9(:monKmo~E"+mb2nK/Yms/W9+r~~w^-|/ODwC6d+OODkd];;kMn9IxkAAA==^#~@%>
						</td>
					</tr>	
					<tr>
						<td>Recip CountryCode:</td>
						<td>
						<input name="RecipCountryCode" type="text" id="RecipCountryCode" value="<%=#@~^QQAAAA==210mwk^VsKDsokV[`rIn^bwZK;xDDX;G[+r~,w^\mdDDolXJ+DYn.b/In5!k.+9bhxgAAA==^#~@%>">
						<%#@~^QwAAAA==21/m"+$EkM+9(:monKmo~E"+mb2ZKExD.zZKNJSP2^7{dYMolXSnOD+Drd"+5Eb.NpBkAAA==^#~@%>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>	
		<tr>
			<td></td>
			<td>
			<input type=submit name="submit" value="Request Signature Proof of Delivery" class="ibtnGrey">
			</td>
		</tr>
	</form>
</table>
<%#@~^QwAAAA==@#@&B&JPG2?:I}eP:C3Pw293oPr~B2;K@#@&dnY,W(Lo+[3XZslkdP{PUGDtkUL@#@&VREAAA==^#~@%>
<!--#include file="AdminFooter.asp"-->