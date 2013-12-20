<%@ LANGUAGE = VBScript.Encode %>
<%#@~^vwEAAA==@#@&BP4kkP0bV~kkP2lMY~G6PnMGN!mY;C.YBPmx~+^Gs:nD1nPmw2sbmlOrKx~N-VGa+9Pl	N,dW^N~4HP3CMVX,(:almD~JS;R,n.W[;1Y/lMO~,kOd,/W;.1+~mK[~~DtPnMW9;mDZCDDPUCs+PmUN,VWTG~lM+,w.W2nMYzPKWPAl.sHPqh2mmO~,JdZ ,ZKwXMkT4Y, T!8O+TZ&R,)V^PDbL4YkPM+d+.-N PIGE,l.n,xWO~mVsWSn9POKP!/+BPmsYDSP9kdOMk4!O+,lx9&GD,D/nVs~mxzPaCDD/~G6Pn.G9E^Y;CMYvkPkWEMm~mKNnPSkO4KEY,OtPhMrOYx,mGxdn	Y~W6~2mDsz,q:2C1Y P:G,mG	YmmY,2m.VHP(:al^OBPw^nlk+P7rdkDPShARnCMVzks2l1Y ^K:R@#@&MpwAAA==^#~@%>
<%#@~^MgAAAA==~alLKbYV'rjtbw2k	o~	bylM[P P2sCrV,IY;DU~dl8+^EPKBEAAA==^#~@%>
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
<%#@~^nj4AAA==~@#@&9b:,;EDHSPM/SP1WUUD+:a@#@&Gk:,rKlT+;E.DnUD~~\m.s^lL(	mWh2^+O+B~!+.H~,/YMr"9~,w^\|kUO}DN.qG@#@&GrhPam7{dY.\Y4W9Hls+S~am\mdDD\+D4KN]w^X~,Z!dYK:nD:DCUklmDrW	qNUOk6kDSP2^7{dYM)m1W;UD1Eh8DSPa^7{dDDt+YDg;:(+.~,w^-|/YM/lMDk./W9+@#@&9kh~am-{kOD:DC^0kxLH!:8+MS,w^7{kYDUtb2:xOb1mG;	Y1!h4D@#@&9r:,w1\m/O.G+dYbUlDkGU;WEUOMX/W9nBP21\|/YMfdYbxCYbWUKK/YmsZKN+B~2m7{kY.SCUTECo/W9+S~am\mdDDJW1C^+/KN~Pam7m/DD9+DlrsUml	d~,wm7mdYMnmorxLPK3nx@#@&fb:~WN+amaWdY9CDlS,W(LsNAaZ^ld/BPG8NrED2ED(Hd9GmBPkD-s39A(p:^uYDwS~w2f3p|Dn/!sD~~w2G2(|j"J~,w^\|/O.ADDK.Hko~,2^\|/DD)mOrKx@#@&@#@&@#@&@#@&v?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?@#@&EPjPzIKl~}1~S})G@#@&EU?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U@#@&@#@&vJz,M2:P}]fAI~qG@#@&21\{kOD}DN.(f{I;;+dOvJrNrb@#@&w^-|/Y.j/dkKU}D[D&f'U+kdkKxcJam)[skx}.NDqGEb@#@&k6P2m-mkY.?d/bWU6MN+.(G'EJ,6"Psxvwm7{kOD}D[+Mq9b@*!PD4+	@#@&i2^\|k	Y6D[nMq9'a^\|/O.}DNn.&f@#@&ij/dbW	`Jamz[:bx6D9+.(GJ#{2m7{k	O6D9+Mq9@#@&n^/n@#@&7w1\mr	Yr.[D(f{21\mkYM?+k/bGx}D[+Mq9@#@&+x9~k6@#@&@#@&vzJPhb!2~HzH3@#@&2mhlLngl:nxrsnNAa|HC	lT+?4kah+	Yd2slrscl/aE@#@&2DMKCo1m:n'EoN36|\l	lLnUtk2hxO/AhmksclkwJ@#@&@#@&BJz~b;K(6g@#@&a^\|/YM)^YbW	PxP.n$En/DcJzmOrKxJb@#@&@#@&BJ&,rKA1,fb:b~)?A@#@&mmVs~Kw+	94v#@#@&@#@&BJz,?3K~P_2~sA92oP6$x2ZP@#@&/nY,G(LoNA6Z^lkdP{PH+SP2^w+NAaZ^l/k@#@&@#@&BJz~M3P,n)Zn)MAP(9,1j\$AIj@#@&KmmVmoqx6W|(f,'~I;;nkY`rKl13lTn(x6W|q9Jb@#@&?n/krW	nC^0lon(	0G{&9,'~U+k/kKxvEw1b[:bxKC13lTnq	0W|(9J*@#@&kWPjnk/rW	Kl13CLqxWG|q9'rE,r],Vx`hl1VlT+(x6Wm(G#@*Z~Y4+x@#@&7w1\|kUYKC13Co(x6W~x,nl^Vmonq	WK{(G@#@&d?/krW	`Ew1b[hbxnm^3mo+&UWW|qGJb'2^7{rxDKl13CLqxWG@#@&nVkn@#@&7am7{k	YhCm0lL+&xWG,'PUn/kkW	KCm0lT+(xWG|q9@#@&nx9PrW@#@&d@#@&Ez&Pw3G2p,Z"2fA1:(bd?@#@&$En.HP',7dr?2d3/K,?4k2:nUDKzwdR!/n.&f~~j4k2:UDKza+kRwm/kAWMNSPUtr2s+xDPXa+/c)^m/kSrmnUk+~J@#@&;!+.z,'P5;DzPL~rs]}H,?tbwsnxDKzw/~E@#@&;!nDHP',5;+MX,[~J	uAI3Pvc`Utr2s+xOPHwn/cr9?4bws+xD#{q#*iEd@#@&dnDPDkx/D\. ZM+mYnr8%mO`r)f}f$ "+mG.9?nYrb@#@&dY,D/{mKUxD+hwc+an1EYc;!+DHb@#@&b0,16K~.kRnW6~Y4+U@#@&dw^-|/ODz^1W;	YgE:(+MxDk`EEk+.(GJ#@#@&dam\|dODt+D+.1;h(+.'Md`rwCdkhW.[r#@#@&i21\mkYM2x7kMGxs+UY{Ddcrbm1n/kSk1nU/J*@#@&+U[,kW@#@&d+DP.d{xWO4bxL@#@&@#@&@#@&EzJP?ASA/K,f)KzPj3:@#@&E~@*@*@*P:C8V/=P2mKC13Co(x6W@#@&$E+.z,'~diEU2JAZ:Pw1nm^3monq	0G CPJ@#@&;!+DH~xP$EDzP'~rs]rt~w1nC^0lon(	0GPr@#@&;;DHP',;!nDHP'Pru3"2Pa^nmm3mLnq	0KR2mKC13Co(x6Wm(G'J~',w^\|r	YKmm0loq	WW,[EPrd@#@&k+Y,./{/+M-nDcZM+CYn6(LnmDcJzf69~RIn^KD[?Or#@#@&/YPM/{^W	xO+sw nX+m!O+v;E.z#@#@&@#@&r0~H}K~Dk +K0~O4+x77@#@&@#@&E&JPJ}rnjn,K_3P"2/qhq3H:PfzPb@#@&B,@*@*@*,/Y~/ndkkGxk~0KD~2M+0rs^N~NmOm@#@&iw1\{kYM6Dborxglhn{D/vEw1nl1VCoq	0G{j4bwoDKhbDYnUDkWUHm:nJ*@#@&d21\|/YMrMrobxJk	+qxM/`r2mhlm0CL+&x6Wm?4ras.Ws)N9DndkFJb@#@&d2m7mkY.}Dbok	SbU+y'./vJ2^hlm0Coqx6Gm?4kas.Wh)9N.+kd r#@#@&iwm-mkY.rMrTkU;kDX'M/vEw1nCm0lLn&x0Km?4kww.G:;kDXE#@#@&iw^\|dYMr.rTkxjOmYnrMKMW-bx1+ZKNxDk`Ew1nC^0lo(x6W{U4rwwDK:jYCOn.W7rx1+/G9+Jb@#@&d2m7mkY.}Dbok	nKdYmV/W9+x.k`Ja^nmm3mLnq	0K{jtr2wDG:hG/Dls/KN+Eb@#@&7w1-|/OMrMkobx;GE	Y.X;W[n{D/vEw1nl1VCoq	0G{j4bwoDKhZKEUOMXZG[Jb@#@&7@#@&v,@*@*@*Pk+D~VKmCV,\C.bl4^n/,0WM~%l7l6ksV~[mYC@#@&7w1\mdDD?nU9+.n4G	+H!:(+D{DkcJamKl13CLqx6G{UtkaPGn4W	+E#@#@&i@#@&B,@*@*@*PdnDPVG^mV~\m.bl8^+kP0KD,%l7lWk^V~[mYli7@#@&dw1-m/DD"+^k2Hm:n'Md`rw^Kmm3CLqU0KmUtraKK1ls+rbdid7did@#@&iwm7m/DDI^rwdk	+q'.dvJ2mhCm0lLn&x0GmUtrw:GzN[M+k/Fr#@#@&dam-{kY.]mkaJk	+ {.d`rw1nCmVCT+(x6G{Utr2:Wb[[M+d/yE*@#@&iw1\{kYM]+1k2ZbYzxM/`r2mhlm0CL+&x6Wm?4raKGZbOXr#@#@&iwm-mkY.I^bwjDlD+rMnMG\bx^+;W[n{D/vEw1nl1VCoq	0G{j4bwPWUOlD+/G9+Jb@#@&d2m7mkY."+1kwhWkOl^ZGN'.dvJw1Kl13lTn(x6W|?4k2PK}rwrb@#@&d2^7{/O."+^ka/KEUDDHZW9+{./vJ2mhl^Vmo+&U0K{?4r2KKZKEUY.zr#@#@&i@#@&Ez&~drrFihPPCA~hb/nbV2P&1w6@#@&d2m7{dOMKDm^3bxog;h4D{Dd`E21nCm0CoqUWK{K.C13rxTH!:8Dr#@#@&da^\|/ODUtr2GlYxDk`Ja^Kl13monqUWK{jtb2wN9CD+Jb@#@&d2m7mkY.;lMDkD;GN'./vJ2^hlm0Coqx6GmsG(;l.DrnMZGNE#@#@&7@#@&Bz&~zf9q:(}1)dPhbIzHAP2"?7@#@&d2^7{/D.?4kw:GKtKx'./cEamKl1VlT+(U6W{j4bwPWh4Kxnr#@#@&dam7m/DDjtbwPGA:lbs'M/`r2^nmm0lL+(U6Wm?4rw:W3hmkVEbi@#@&dUnk/rKxvJw1b9hk	j]SgWOr6kmmOkKx2tCrVzN9Dn/dE*'2m7m/DDj4bwKG3slrVi7@#@&7am7{/DDhCm0lL+d+ULDt'Md`rwmhC^3moqU0Gmhl^3mL+d+ULDtJb@#@&d2m7mkY.hl13lT+qrNDtxDk`E21nl1VlT+q	WG{hl13Con	bNOtrb@#@&d2^7{/O.hl^3mLCnbo4Y'M/vEw1nCm0lLn&x0Kmnmm3mLnCkTtOJb@#@&d2m7m/DDKC13lLnq+ro4O{DdvJamnmm0CoqU0K{KC13lTnko4OE#@#@&iw^\mdDDKl1VlT+(UkEDn[jlsExM/crw1nl13mL+&xWW|nC^0lo(xkED[#l^EJb@#@&7@#@&nx9~k6@#@&dYP.d{xGY4r	o@#@&@#@&BzJPU3K,I3p`q]3GP.z]qzASAj@#@&am7{dY.\Y4W9Hls+~x,Js9pA:Ck^Jm4n^I;E/DE@#@&w^\|/O.t+Y4GN"+w^z~',Jwfp2hCbVJl(nV"+2sHJ@#@&/!/OWsnMK.mxklmDkKUq9+UYb0rnMP',EnMWN!^OZmDD{3:Cr^InY!.xrd@#@&am\mdDDjtb2s+UDb1mW!xDHEs4nD{w^-|/YM)m1WE	OHEs4D~B&&,rAx.BkP)^1WEUO,1;:(nM@#@&@#@&@#@&B?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U@#@&B,21G),KlT+~SKl[@#@&BU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U@#@&@#@&@#@&@#@&veCeeCMMeCeCeMeMMCeMeCMeCeMMCeeMMCeMeCMCeMCeCeeCeCMeCeMeCeMMCeeCMeCeeCMMeCeC@#@&B~j:b]K=~n}?P~~bZF@#@&BMeCMCeMCeCeeCeCMeCeMeCeMMCeeCMeCeeCMMeCeCeMeMMCeMeCMeCeMMCeeMMCeMeCMCeMCeCe@#@&k6~D;;+kY WKD:vE/!4:bOE#@!@*rJ~Y4n	@#@&@#@&7B?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U@#@&iv,MnDPmVV,W6~Y4+~D;;rM+N,rx6WDsCOkKxc@#@&dv=?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U=7@#@&d@#@&iB&z,!xnMk1P+MDK.P6W.PalLn@#@&da^\|/YM!nxDbmKlLnAD.WM~',J)O,V+CdDPGx~M+5!kM+N,0bnV9PAlkPnhaYXc~J@#@&d@#@&7w1/|.CVr[mYnKaYwkns9dJi]d2awb.mYrKxGlYJB~YMEn~,!@#@&iwmkm.mVk9CO+:+XYokns9dEj"J1KYrWbmlOrKx3Hmr^b[9D//r~,OD!+SPZ@#@&7am/|#l^kNmOnK6Dsr+s[iJ9DK2W60Pza+JS~6ls/S,!@#@&dam/|.msk9lO+:+aOwk+^[dr?+M-rmJBPOD;nBPT@#@&7w1/m#mVk[CD+P+XOwkn^NiJnmm0CobxLJBPO.!+~,T@#@&dw1dm.mVbNCYnP6OsbnV9dE(D+:9nkm.kaObWUr~,YD!+B~%Z@#@&d@#@&7vam/|#l^kNmOnK6Dsr+s[iJiIdHWDkWr1lYrG	2\lbst+dklT+JBP6CVk+SPZ@#@&7am/|#l^kNmOnK6Dsr+s[iJ]+krNxOrmVnr^0E2JB~6lsk+BP!@#@&i2mk{#l^k[CD+KaYwk+^[7J"+kk[+UOblsfsk7+.zr~PWC^/n~,Tid@#@&dEwmk{jCVbNCYKnaDsksNiJq	dOD!mDkGxdEBPWl^d+BPT@#@&dB2^k{#l^r9lOK6YwksNiJ]+1k2rxYdGmmYkKUHEs4DE~~WmVd+B~!i@#@&7Ewmdmjlsk9CD+P6DskV97J"H)1!:8nMdJB~0mV/S~!i@#@&d2mdmjlsk9CYKnaDskns9dEnmzKDPHwJ~,0ms/~~!@#@&721/{jCVbNlDnP+XYwknV[7rnCXK.b1mG;	Y1;h(+.JB~6lsk+BP!@#@&i2mk{#l^k[CD+KaYwk+^[7JhlHW.ZG;	Y.X;GNJS~6lVdnBPT@#@&7amd|.mVk9lDnK6Osb+s[iJ?bLxmYEMn6wDkKxE~~WmVd+B~!i@#@&@#@&dkW~U+d/bG	`EamzN:bx"n/bNnxDkCshkm0;wr#'rE~Y4+	@#@&d7j/dkKU`rw^)9:kU]/rNUDkC^nbm3!wrb'Z@#@&dx[~b0d@#@&db0PUnd/bW	`Ew^)9:rx"n/bNnUDkls9Vr\.HJb{JrPY4+	@#@&idj+k/rG	`Ja^b9:k	]n/bNxOkCsG+sk7nDHJbxZ@#@&7n	N~k67@#@&7EU?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U@#@&iB~/4+m0~0KDPjCsk9lDkGx~3MDGDk PGW~UKYP2.Kmn+9~b0~DtD+,lMnPD.WM/ @#@&dB?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?@#@&d(0~21\mk	O2MD@*T,KtnU@#@&7dMnkwG	/RDNb.+1Y~w1nCL1lsnPLPJQhdo{J,[~w^-|/ODVnxDr^hlon3MDGD@#@&i2sk+@#@&did@#@&diB=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?@#@&i7B,AEbs[P}EMPPDCUkl^YbGxc@#@&7iBU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U@#@&idG(Lw+NA6;slk/ 1hp\dKDmU/mmYbGUPam7{dY.\Y4W9Hls+S~am\mdDD)m1G!xOgEs4+M~,2m7{dYMHnOD1!h4D~,2^\|/DD/l..b+.ZK[+BP/;kYWhnMK.l	dmmObW	qNxDr0b+.@#@&d77@#@&di7dK4Lwn[2XZ^ld/ 	MkO+UrxTVnKmD+UO,JiId3XwrMlDkW	fmO+r~~?/drKx`r2mzN:bUiId2XwrDCObWUfmO+r#77idd77@#@&7di7K4%w+926;Vmd/c.kD+jr	oVKlM+xD~Ej"SgWOkWr1lOkKU2 HCr^bN[./dJB~U+dkkKx`rw1)NskUj"SHGDk0b^lDkW	3\lbVzN[DndkJb@#@&7didG8Ns+[3XZslkdc.bY?k	o^nnmDnxDPEi"S1KOk6kmmOrW	2 HCks\/dlTnJBPjnk/kGUvJ2mz[skU`Id1WDk6rmmYrW	2\CbVHd/mo+rb@#@&ididG4%oN36;slk/ 	MkYnjbxLVKmDn	Y,JHD14l	YKtKxnH!:4.JBPw1-m/DDU+UNn.htGxHEs4n.@#@&d77iW8Lwn92a;Vm//cMrY?rxTVnKmD+	OPrI+kr[+	Yblsnr^0E2JB~?/drKx`E21b[:bU"+dbNxYbl^Kk13;wr#@#@&iddiG4Ns+93aZ^lk/ .rD+jk	LVnC.xY~EGDGwKW6Kza+r~PU+kdkKxcJam)[skxG.WaW06PzwJ*@#@&d77iW8LwnNA6/sm// 	MkO+Ur	osnmD+	Y,E?D-k1+ES,?+kdkKx`r2^b9:bxj+.-bmnJ*@#@&id77K4Lon92aZ^Ck/ qDbY+Uk	LVnCDxO~rnl1VlTkxTESPU+k/rWUcrw^b9hk	nC^0lorUTJb@#@&7id7@#@&iddiW(%sN36;VCdkRMrY?k	Ls+hlM+UY~Eq+ro4Oj	kOdr~PEJ~?E@#@&7id7K4Ns+92X/Vm/dRqDrO?k	LVnlMnUY,J;E.DnU1X/W9nJBPEiUfJ@#@&id7d@#@&id7iW(LsNAaZ^ld/c.rD+nm.+	YPr6.kTk	JSPEEid7di7d@#@&77iddG8NsnNAa;VCk/cDbYKlM+UY,J/G	Yl1OJBPJr77didid@#@&77id7dK8Lw+[3XZVCdkR)N9HhHKNPJh+MdW	1C:JS~am\|dYMrDbLrxgls+@#@&77id7dK8Lw+[3XZVCdkR)N9HhHKNPJhtKU+gEh4DES,wm7m/DD?U[+Mn4WU+H;s4nDi7did@#@&idd77K4%s[A6/^lk/RqDbO+hl.+	Y~E;WxDCmDJ~,E&Ji@#@&d7d77K4%s[2XZsCk/R	.bYnnm.xO,JzNNM+kdJBPEJ@#@&77iddiG4Ns+93aZ^lk/ b[[g+A1K[+,JJr	+FES,w^\|dDD6MkTkxdk	nF@#@&7did77K4LwnNA6Z^Cd/cb9NH+AHKNnPrJk	++EBPw^-|/OD}.bor	Sbx+y@#@&7did7dK4%oN2X/Vm//c)[Ng+S1GNn~rZrYHE~,w^-|/Y.6MkLk	/bYz@#@&iddidiG4NsnNA6/sm//c)N91+SHGNPr?OlOn}DKDK-k	mn/KN+ES,w^\|dDD6MkTkxUYmO+}DKDK\rU1+ZK[+@#@&di77diW(Lo+[3XZslkdRzN[Hh1G[PEnKdDls;W9+JBPa^\|/OD}DrLbxnKdYmVZK[n@#@&did7d7G(Lo+936;VCdkRb[[g+A1K[PE;W!xYMX;GNJSPam-mkYD}.kTkx;G;xDDHZGNn@#@&d7di7W(Lon926/sm/dRq.bYnhlM+xDPr)N9Dn/kJS~rzJi7di@#@&i77dK4NsnN3a;VC/k MkOnhlDnUDPErMrTkUr~,Jzr@#@&@#@&id7dK4%oN2X/Vm//c	.kD+hl.+UO,J9+kOk	lOrKxJS~rJ7di7id7@#@&iddidK8Lw+[2XZsCk/Rq.kD+nm.nxDPrZGxOC1YE~,EJ@#@&77idd7G(Lo+93XZsm/kRb9NgnhgW[+,JKnM/W	Hls+JB~2m7{kY.In^bwHlsn@#@&d77iddG8NsnNAa;VCk/cbN91A1KNnPrn4G	+1!h4DJB~2m7{kY.?4raKGn4Gxd77idd77i@#@&di7idG(Lw+NA6;slk/ MkOnhlDUY,JZKUOl1Yr~~J&Ei@#@&di7diW8%w+N3a;VC/k qDrD+hlDxD~JzN[D/dEBPJr@#@&iddi77W(Lw+[2a/^ld/c)N91nAgWNn~rSrxqr~~am7{/DD"nmbwJk	+q@#@&ddi7diW4NonNA6;VC/d zN[1A1KNn~rSkUnyJSPa^7{dDD"+mbwdrx @#@&id77idW(%sN2X/slk/cb[NHnS1GN~J;kOzr~P2^7{dYM]mraZbYX@#@&i7did7W(Lon926;slk/Rz[[1hgW[+~EUYCY6DhDG-bxmn/KNnJB~am-|/DDImb2?DlO+}DK.K\k	^+;WN@#@&didid7W8%w+[2X/Vm/d zNNHnS1GN~rnGkYmVZKNE~,w^\|/O."+mb2nK/Yms/W9+@#@&7d77idG4No+92a/^l/d zN[1AgW[PrZW!xD.X;W[+r~~21\{kOD"+mb2/W!xDDzZG[@#@&di7diW8%w+N3a;VC/k qDrD+hlDxD~JzN[D/dEBPJJEdidd@#@&7didK4%sn[A6/Vmd/c.rD+nC.xOPr9/ObxmYkKxrSPrzEdid@#@&iddi@#@&iddi@#@&didiW8Lon92aZ^C/kR	.bY+KCM+UY,Ehlzs+	YJBPrE@#@&d7didG8Ns+936;Vlkd b9Ng+A1G[PEnmzWMKz2J~~EU2HfA]r@#@&ididdb0,j+k/rW	`E21bNsrxhlXK.)m1W!xO1;h(+.J*~@!@*PEE,YtnU@#@&7di7idG(Lw+NA6;slk/ MkOnhlDUY,JnmzGDr~,JE@#@&7id7di7W(Lon926/sm/dRz[91nS1KN+,Jz^mKEUYgEh8DJB~?//bGU`rw1b[:rUhlzWM)m1W;UD1Eh8DE#@#@&id7ididW(LwnNA6/Vm/d zNNgnhgWN~EZKE	Y.X/G9+E~,j+k/rG	`J2^zNhk	KmXGMZKExDDH/W9+E#@#@&77iddiG4Ns+93aZ^lk/ .rD+KlMnxDPEKmXW.EBPEzr@#@&d7idi+x9PbW@#@&d7diW8%w+NAaZ^l/k 	DbYnCDnUDPEnmz:xOEBPJ&Eid7di7id7ididd@#@&@#@&did7W(Lon926;slk/Rq.rYnmDnxO~r?2+1rl^?n.7kmndr~~Jr7id7idi@#@&idi7dK4%sN3a;VlkdRqDkDnKlM+	Y~J3\mks1KOk6k^CDkWUEBPEJ@#@&id7ididW(LwnNA6/Vm/d qDkDnnmD+	O~JUtbw2+.EBPEJ@#@&did77iddG8NsnNAa;VCk/cbN91A1KNnPr?4rabV.Yr~P8@#@&didid7d7G(Lo+936;VCdkRb[[g+A1K[PEG+^k\DHHWDkWk1lOrKxJB~F@#@&di77didiW8Lon92aZ^C/kR)[91+AHKNnPr3XmnaYbWxgWDr0bmCYbWUEBPF@#@&diddi77dK4NsnN3a;VC/k b9NHnS1W[n,JoWMhmYEBPrCKtSr@#@&id7did7G(Ls[2XZVmddRqDbYnnC.xOPrjtbw2nMJ~~EJJ@#@&i7id7idK4Lw+936;VC/kR	.bY+hCDxY,E]+1kaknxOEBPEJ@#@&did77iddG8NsnNAa;VCk/cbN91A1KNnPr?4rabV.Yr~P8@#@&didid7d7G(Lo+936;VCdkRb[[g+A1K[PEG+^k\DHHWDkWk1lOrKxJB~F@#@&di77didiW8Lon92aZ^C/kR)[91+AHKNnPr3XmnaYbWxgWDr0bmCYbWUEBPF@#@&diddi77dK4NsnN3a;VC/k b9NHnS1W[n,JoWMhmYEBPrCKtSr@#@&id7did7G(Ls[2XZVmddRqDbYnnC.xOPr]+1k2rxYES,J&J@#@&id7ididBJz,fD9P2lMYz~	WYbWk1lYbGU/@#@&id7d77iW8LwnNA6/sm// 	MkO+hCM+UDPrrY4+ME~,JE@#@&d77iddi7W(Ls[36;Vm/dR)[91nhgGNPE3tlks)9N.+kdr~~U+k/kKxvEw1b[:bxi]d1WDr0bmlDrGxAHmksb[[M+d/rbP@#@&77idd77iW8Lwn92a;Vm//cb9[1hHW9+~EUtka)VDYrS~F@#@&id7d77idG4No+92a/^l/d zN[1AgW[Prf+^k7nDH1GYb0r^mYkKUJBPF@#@&7didid7dG8NsnNAaZ^lddcbN[HhHW9n,J3XmwYbW	HWDkWk1lOrKxJB~F@#@&di77didiW8Lon92aZ^C/kR)[91+AHKNnProKDhmYr~PrC:\Sr@#@&did77idW(%sN2X/slk/c.kOnhl.+	OPrrO4DJS~rzE@#@&7id7iW(LsNAaZ^ld/c.rD+nm.+	YPr3\lbVgWOkWr1lOkKUJBPE&r@#@&77id7@#@&7id7iW(LsNAaZ^ld/cb[[g+hgGNPJ"ndk9+	Yrls9Vr\.Xr~~j//rG	`Ew1)9:r	I/k9+	OkmV9+^k-nMXJ*@#@&iddi77didid7d@#@&id7dK8Lw+[3XZVCdkR	DbOnCM+	YPr?anmbls?D-r1+/rSPrzJ@#@&7didid7d77i@#@&di7dK4%oN2a/^ld/c)9NHhgWNPr$VKmV?4k2hxYGCYmJ~,q7didi@#@&@#@&7id7BK8Lw+[3XZVCdkR	DbOnCM+	YPrIt)JBPEJid77iddi@#@&iddiv7W(Lw+[2a/^ld/c)N91nAgWNn~r1;:(nMJS,?//bW	cJam)NskU]tb1!h4DJ*77didid7d7@#@&d7divW(Lon926/sm/dRq.bYnhlM+xDPr]HzJSPrzE@#@&ddi7@#@&ddi7G4NsN36/sm/dRq.kD+KCM+xO~rnCm0CT+EBPrJdidi7did7did77id@#@&7diddK8%sNA6/VCdkR)N9H+S1G[PJ9n1VCD[jls!+r~Pam7m/DDKl13CLqxk;DN.ms;+idid7d77@#@&7di7db0~j//rG	`Ew1)9:r	?D\bmE#,'~JIri]hbZn)M&1Mr~Otx@#@&7d77idG4No+92a/^l/d qDrYKmDn	Y,Jfb:U/bWU/r~~Er@#@&i7diddiG8Lw+92aZsCk/ b9[1hHG9+PEJxLY4EBP21\|/YMnm^3monSxLO4@#@&i7diddiG8Lw+92aZsCk/ b9[1hHG9+PE	bNOtrS,w^7{kYDhl1VlT+	k9Y4@#@&ddi7didW(%o+92XZslddcb[NgnhgW[n,JCnrTtOJB~am-|/DDnmm0CoCnkTtO@#@&ddi7diW4NonNA6;VC/d qDrYKlM+UO,JfrhxdkKUkJS,JJJd@#@&i7didnx9PrWiddi7did@#@&77didK4%sn[A6/Vmd/cb[[g+hHG9+~Jqnbo4DJBPsKDsCYgEh4Dc21\{kODhlm0CL+q+bo4YSq*@#@&di7di@#@&7idd7vK4%s[A6/^lk/RqDbO+hl.+	Y~E"+0.+	m+&UWWr~,JE@#@&7id7dE7W(Lon926/sm/dRz[91nS1KN+,J;;/DWh+MInWD+	^+r~PUnd/bW	`Ew^)9:rx;;/DWhnMI+WnM+UmE*@#@&ididdEdK8Lw+[2XZsCk/Rz[Ng+hgG[+,JhrHEh8DE~,j+k/rG	`J2^zNhk	/!/OK:Dn}1!h4DE#@#@&77iddE7W(Ls[36;Vm/dR)[91nhgGNPE(	\Wr^1;:(nMJS,?//bW	cJam)NskU/!/YKh+Mqx7Grm1!:8+.E*@#@&di7diBG8Ns+[3XZslkdc.bYnlM+	OPrIn0DnU1+q	WWr~Pr&E@#@&did7d@#@&id7divW(Lon926/sm/dRz[91nS1KN+,JUro	lOEM+62DkW	E~,JfAJ(.AIqqPC6i:?(Mg)K`I3E,Bz&~G2Jqj3"(:C}jKUqVHb:j]2BP(HGqIA/KBPf&]3Z:~,b9jJP@#@&7di7dK4%oN2a/^ld/c)9NHhgWNPr(Y:9+km.raYkKUJBP?ddkKxvJ2m)[skUqDn:G+d^MkwOrKxE#i7id7i@#@&ddidK8Lw+[2XZsCk/Rq.kD+nm.nxDPrnCmVCT+E~,Ezr@#@&7@#@&d7G(Lo+93XZsm/kR2	No\S:DCxkl^ObWx,2m7{/D.\+DtKNHlhni@#@&di@#@&idv&JPn.r	Y~W!O,W;MP	+h^X,WWM:nN,Dn5!+/D~6sV@#@&77BM+kwGxdnch.kDnP6+[nX{wGdDNCYm@#@&d7ED/wKxknRx[@#@&d7@#@&ddE=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?=@#@&idEPj+U[,r;D,PDmxdC1YkGUc@#@&div?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?U@#@&di^l^V~W(Lon926;slk/RUnUNoHdIn;;nkYc0[+X{2GkYNCOm~~w1-|/OM2	\kMW	h+	Yb@#@&d7@#@&ddE&z,nDbUOPKEDPGE.~M+dwKU/@#@&7iBDndaWU/ SDrD+,s2G2omD/;VD@#@&7iBDdwKx/ nx9@#@&d7@#@&7iB=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U@#@&idB~JKl[P};MP]/aWxk+c@#@&idvU?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?U@#@&7d1lV^~G4NsN36/sm/dRdGl9(\J"+/;sD/csA9A(mM+kEVD#@#@&di@#@&di@#@&7iBU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU@#@&7dEP;tnmV~6W.P.DKDd~6DWh~w+[2X @#@&7iB?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=did7@#@&d7^mVV,G4Ns+93aZ^lk/ (\J"+dwKU/.n.b0Xc3MDKlTnglh#@#@&di@#@&7d@#@&7dEU==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?@#@&idB,]nNbDmOPArDt~l,\+k/CLPr]~1Whw^nD+~kWs+PDlkVR@#@&7dEU==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?@#@&idk6~Hr:P^+U`2^7{dYM3DMW.\ko#@*T,Y4+	@#@&d7@#@&id@#@&di7B?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=@#@&d77EP?OP}ED,]n/aW	/nP9CDl~YK~SKmCsc@#@&77iB=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?@#@&d7iB,b\mk^C4^+~HY4G9/PSrV^P/C.m4P!xskhrD+[P^n\Vd~K0PHG9+dP(z,/nalMlYbxT~xKNn/,hrO4Pl,EzrR@#@&77dEP8#~InC9In/aGxk+KCM+xO@#@&d7dE~y#~"+mNI/aGxk+HW9+@#@&idd@#@&didBJ&~CAbG2]@#@&7id2m7m/DD/;kYWhnMK.l	dmmObW	qNxDr0b+.P{PG8Ns+936;Vlkd Il9In/2G	/n1K[+vJ&&"+wsz_+CN.r~~rZ!/YK:.KMlU/mmOrKxq9nxDk0bn.J*d@#@&7d7@#@&d7dE&z,2]]}I@#@&7id2m7mkY.ADMWD;W9nP{PG4Nsn[A6Z^C/kRIC[I/aWU/nHKNn`r&zAD.GMJ~~E;W[+rb@#@&7idam\|/D.2MDGDt+ddmo+,xPK4Lwn[2XZ^ld/ ]l[IdwKxdngWNncrz&2M.KDEBPrH+k/mL+r#@#@&id7@#@&ddi2m7{/D.iIdP{PG4%oN36;slk/ ]lN]nkwGxknhl.xD`w1\|dYMHnY4W[]wVHSPrjIdEbdidid7@#@&7id2m7m/DDidDq9~{PG4NoN3XZ^l/kR"nl9In/aWUdnlMnxD`w1-m/DDt+OtG["+2VHSPrjdnMqfEb@#@&7di21\mkYMnlk/SGD9PxPK4%oN2X/Vm//c]nl9I/2WUdnCDUYvw^-|/Y.\Y4W9]wsH~,Jnm/kAWMNE#@#@&77iwm7m/DDI5;+kY:kh+jOm:2P{~W(Lon926/sm/dR"nmN]/aWxk+hCDxO`am-mkYDtnY4WN"n2VH~,J]+5;/OKbh+UYChaJ#@#@&id7@#@&7idvJz,nb;|z!2@#@&7diw^-|/YMPDmm3bUL1!:(+.Px~K4%s[2XZsCk/R]nmN]+k2Kxd1KN+vJJ&nmmVlT+ES,JKMCm0kxTH;:(+MJb@#@&7id2m7m/DDoGM:q9~{PG4NoN3XZ^l/kR"nl9In/aWUd1W9n`rzzhC^3moJSPEoKDhqGE#@#@&77i@#@&77iw^\|dDDjbo	lY!D6wDkGx,'~G(Ls[2XZVmddR"+mN]+d2Kxd+hCDxOcam\mdDD\+D4KN]w^X~,JUro	lOEM+62DkW	E#@#@&di7@#@&idiJuETAA==^#~@%>
			<table class="pcCPcontent">
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr>
					<th colspan="2">FedEx<sup>&reg;</sup> Email Label Confirmation</th>
				</tr>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>		
				<tr> 
					<td>Email Label Link:  </td>
					<td><a href="<%=#@~^CgAAAA==21\mkYMjId9AMAAA==^#~@%>" target="_blank">Click Here to View</a>. This link should be saved; it will only be valid for 10 days.</td>
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
		<th colspan="2">FedEx<sup>&reg;</sup> Email Return Label Request</th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
			<p>
			A FedEx Email/Online Label is sent directly to the return shipper without exposing the FedEx customers account information. Returns shippers print labels from their own
			printer using an application on fedex.com, place the label on the package, and drop-off or request pickup.
			<br />
			<br />
			<span class="pcCPnotes">This form contains only standard return label options. For additional return label options, please visit fedex.com.</span>
			</p>
		</td>
	</tr>
</table>

<table class="pcCPcontent">

	<form name="form1" action="<%=#@~^CgAAAA==21nCT+gl:0QMAAA==^#~@%>" method="post" class="pcForms">
	<input name="PackageInfo_ID" type="hidden" value="<%=#@~^EgAAAA==21\mbxDnl13mL+&xWWKwcAAA==^#~@%>">
	<input name="id" type="hidden" value="<%=#@~^DgAAAA==21\mbxDrD9+M(ffAUAAA==^#~@%>">
	<%#@~^BQIAAA==@#@&d[D?4kwa+99lD+xfmYn)9N`r[JBF!B[CY`*#@#@&7r6PjpdmsKDhCD'JqE,Y4+	@#@&d79YUtkaw[fmYn'vNCzvNYU4kaw+99CY#LJ&J'hKxOtv[YUtr2a+N9CD+b[r&r[zlM`ND?4rwa+[fmYnb*@#@&inVk+@#@&77ND?4k2wn[GlO+{c:KxO4vNYj4bw2+99mYn*[rzJLNmz`9Yjtbw2n9flDn#LJzr'z+mDvNO?4rawnNGCY#b@#@&d+U[,kW@#@&76EU1YbWx,wm[`Dtn\mV;n*@#@&i76{V+	cOt\mV;+b@#@&d7k6~6,'~q,YtnU@#@&7di2mNxr!r[Y4+7CV!+@#@&idnsk+@#@&7diwl9xOt\mV;+@#@&idnx9~k6@#@&7xN~W!x^YbG	@#@&iND?tbwanNGlO+{`znmD`9O?4kwan[fmY#'JRELwCNvhW	Y4c9Y?4rawnNGCD+b*[rOJLwm[`9lz`9Yj4bww[fmY+*bb@#@&deJYAAA==^#~@%>
	<input name="URLExpirationDate" type="hidden" id="URLExpirationDate" value="<%=#@~^DQAAAA==[D?4bwa+NGlDnIwUAAA==^#~@%>">
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>		
		<tr>
			<th colspan="2">FedEx Email Return Label Request Options</th>
		</tr>		
		<tr>
			<td colspan="2" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td>Carrier:</td>
			<td><%=#@~^EgAAAA==21\mkYMZlMDbnD;W[+RAcAAA==^#~@%></td>
		</tr>
		<%#@~^MQAAAA==~@#@&7ik6Pw1\|dYMZCDMkn.;WN~',JsGp3J,Y4+UP@#@&idfAwAAA==^#~@%>
		<tr>
			<td>Drop off Type:</td>
			<td>
				<select name="DropoffType" id="DropoffType">
				<option value="REQUESTCOURIER" <%=#@~^MAAAAA==210mU+^+mDraOkKxcJGDG2K00:zwJ~r]3p`2UK/ri]&2]J*KhAAAA==^#~@%>>Courier Pickup</option>
				<option value="REGULARPICKUP" <%=#@~^LwAAAA==210mU+^+mDraOkKxcJGDG2K00:zwJ~r]3M`SzIKq/F`nE#xg8AAA==^#~@%>>Regular Pickup</option>
				<option value="DROPBOX" <%=#@~^KQAAAA==210mU+^+mDraOkKxcJGDG2K00:zwJ~r9]rhA}(E#Bg4AAA==^#~@%>>FedEx Express Drop Box</option>
				<option value="BUSINESSSERVICECENTER" <%=#@~^NwAAAA==210mU+^+mDraOkKxcJGDG2K00:zwJ~r$i?&1A?j?3]jq/2;31:2]E*JhIAAA==^#~@%>>Business Service Center</option>
				<option value="STATION" <%=#@~^KQAAAA==210mU+^+mDraOkKxcJGDG2K00:zwJ~rjPb:q}1E#Cg4AAA==^#~@%>>Station</option>
				</select>		
			</td>
		</tr>
		<%#@~^CAAAAA==~x[,k6PZgIAAA==^#~@%>
		<tr>
			<td>Service Type:</td>
			<td>
					<select name="Service<%=#@~^AQAAAA==VawAAAA==^#~@%>" id="Service<%=#@~^AQAAAA==VawAAAA==^#~@%>">
					<option value="FIRSTOVERNIGHT" <%=#@~^LgAAAA==210mU+^+mDraOkKxcJU+.-bm+r'3BJs&]jK}.AIHq!u:JbFg8AAA==^#~@%>>FedEx First Overnight&reg;</option>
					<option value="PRIORITYOVERNIGHT" <%=#@~^MQAAAA==210mU+^+mDraOkKxcJU+.-bm+r'3BJn"(6I&KIr#2]H&MuKrbEBAAAA==^#~@%>>FedEx Priority Overnight&reg;</option>
					<option value="STANDARDOVERNIGHT" <%=#@~^MQAAAA==210mU+^+mDraOkKxcJU+.-bm+r'3BJ?:)HfzIGr#2]H&MuKrb3w8AAA==^#~@%>>FedEx Standard Overnight&reg;</option>									
					<option value="FEDEX2DAY" <%=#@~^KQAAAA==210mU+^+mDraOkKxcJU+.-bm+r'3BJsA93(yfz5E#VA0AAA==^#~@%>>FedEx 2Day&reg;</option>
					<option value="FEDEXEXPRESSSAVER" <%=#@~^MQAAAA==210mU+^+mDraOkKxcJU+.-bm+r'3BJsA93(A(hI3?jjz.3Irb7w8AAA==^#~@%>>FedEx Express Saver&reg;</option>
					<option value="FEDEXGROUND" <%=#@~^KwAAAA==210mU+^+mDraOkKxcJU+.-bm+r'3BJsA93(VI}jHfEbEw4AAA==^#~@%>>FedEx Ground&reg;</option>
					<option value="GROUNDHOMEDELIVERY" <%=#@~^MgAAAA==210mU+^+mDraOkKxcJU+.-bm+r'3BJM"6i1GC}H3f3J&.3IIE#NBAAAA==^#~@%>>FedEx Home Delivery&reg;</option>
					<option value="INTERNATIONALFIRST" <%=#@~^MgAAAA==210mU+^+mDraOkKxcJU+.-bm+r'3BJqgP3Igb:q61)Jwq]?:E#OBAAAA==^#~@%>>FedEx International First&reg;</option>
					<option value="INTERNATIONALPRIORITY" <%=#@~^NQAAAA==210mU+^+mDraOkKxcJU+.-bm+r'3BJqgP3Igb:q61)JhI(r"(KIJbMhEAAA==^#~@%>>FedEx International Priority&reg;</option>									
					<option value="INTERNATIONALECONOMY" <%=#@~^NAAAAA==210mU+^+mDraOkKxcJU+.-bm+r'3BJqgP3Igb:q61)JAZ61}\5r#yhAAAA==^#~@%>>FedEx International Economy&reg; </option>
					<option value="INTERNATIONALPRIORITYFREIGHT" <%=#@~^PAAAAA==210mU+^+mDraOkKxcJU+.-bm+r'3BJqgP3Igb:q61)JhI(r"(KIs]3&MCPE*OxMAAA==^#~@%>>FedEx International Priority&reg; Freight</option>									
					<option value="INTERNATIONALECONOMYFREIGHT" <%=#@~^OwAAAA==210mU+^+mDraOkKxcJU+.-bm+r'3BJqgP3Igb:q61)JAZ61}\5wI3(VCKEb0xIAAA==^#~@%>>FedEx International Economy&reg; Freight</option>
					<option value="FEDEX1DAYFREIGHT" <%=#@~^MAAAAA==210mU+^+mDraOkKxcJU+.-bm+r'3BJsA93(8fz5oI3(VCPJ*XA8AAA==^#~@%>>FedEx 1Day&reg; Freight</option>
					<option value="FEDEX2DAYFREIGHT" <%=#@~^MAAAAA==210mU+^+mDraOkKxcJU+.-bm+r'3BJsA93(yfz5oI3(VCPJ*XQ8AAA==^#~@%>>FedEx 2Day&reg; Freight</option>									
					<option value="FEDEX3DAYFREIGHT" <%=#@~^MAAAAA==210mU+^+mDraOkKxcJU+.-bm+r'3BJsA93(2fz5oI3(VCPJ*Xg8AAA==^#~@%>>FedEx 3Day&reg; Freight</option>
					<option value="EUROPEFIRSTINTERNATIONALPRIORITY" <%=#@~^QAAAAA==210mU+^+mDraOkKxcJU+.-bm+r'3BJ2`]6nAs&IjK(H:2]1zPq}1)JhIq6]&KeJ*ihQAAA==^#~@%>>FedEx Europe First International Priority</option>
					</select>
					<%#@~^JgAAAA==21/m"+$EkM+9(:monKmo~EU+D7rmJ[0S~YMEtw0AAA==^#~@%>
			</td>
		</tr>
		<tr>
			<td>Package Type:</td>
			<td>
					<select name="Packaging<%=#@~^AQAAAA==VawAAAA==^#~@%>" id="Packaging<%=#@~^AQAAAA==VawAAAA==^#~@%>">
					<option value="FEDEXENVELOPE" <%=#@~^LwAAAA==210mU+^+mDraOkKxcJhl^Vmok	LJL3~ro3fA(A1#2J6h2E#Vg8AAA==^#~@%>>FedEx&reg; Envelope</option>
					<option value="FEDEXPAK" <%=#@~^KgAAAA==210mU+^+mDraOkKxcJhl^Vmok	LJL3~ro3fA(hbFJb1A0AAA==^#~@%>>FedEx&reg; Pak</option>									
					<option value="FEDEXBOX" <%=#@~^KgAAAA==210mU+^+mDraOkKxcJhl^Vmok	LJL3~ro3fA(~rpJb4Q0AAA==^#~@%>>FedEx&reg; Box</option>
					<option value="FEDEXTUBE" <%=#@~^KwAAAA==210mU+^+mDraOkKxcJhl^Vmok	LJL3~ro3fA(:j$2EbKA4AAA==^#~@%>>FedEx&reg; Tube</option>
					<option value="FEDEX10KGBOX" <%=#@~^LgAAAA==210mU+^+mDraOkKxcJhl^Vmok	LJL3~ro3fA(8!FM$6oJb1A4AAA==^#~@%>>FedEx&reg; 10kg Box</option>
					<option value="FEDEX25KGBOX" <%=#@~^LgAAAA==210mU+^+mDraOkKxcJhl^Vmok	LJL3~ro3fA(y*FM$6oJb2g4AAA==^#~@%>>FedEx&reg; 25kg Box</option>
					<option value="YOURPACKAGING" <%=#@~^LwAAAA==210mU+^+mDraOkKxcJhl^Vmok	LJL3~re6j"nzZFb!(gME#YA8AAA==^#~@%>>Customer Package</option>
					</select>
					<%#@~^KAAAAA==21/m"+$EkM+9(:monKmo~Ehlm0Cobxor'V~,YMEnaw4AAA==^#~@%>
			</td>
		</tr>
		<tr>
			<td>URL Notification EMail Address:</td>
			<td>
			<input name="URLNotificationEMailAddress" type="text" id="URLNotificationEMailAddress" value="<%=#@~^NgAAAA==210mwk^VsKDsokV[`rj]JgWYbWk1lYbGU2tlbV)N[./dJB~YMEnboBMAAA==^#~@%>">
			<%#@~^OAAAAA==21/m"+$EkM+9(:monKmo~E`ISgGYb0k1COkKxAHCks)9N.+kdJBPO.!+vRQAAA==^#~@%>
			</td>
		</tr>
		
		<tr>
			<td>Item Description:</td>
			<td>
			<input name="ItemDescription" type="text" id="ItemDescription" value="<%=#@~^KgAAAA==210mwk^VsKDsokV[`rqOnsf+k^DbwYbGUJBPDD;+bSw8AAA==^#~@%>">
			<%#@~^LAAAAA==21/m"+$EkM+9(:monKmo~E&Y+s9+kmDb2OkKxr~~Y.;aBAAAA==^#~@%>
			</td>
		</tr> 		
		<tr>
			<td>Residential Pickup:</td>
			<td>
			<input type="checkbox" name="ResidentialPickup" value="1" class="clearBorder" <%=#@~^KQAAAA==210m;tm3}wDrW	`EI/r[xYbCVhkm0;2JBPrFE#JQ4AAA==^#~@%>>
			</td>
		</tr>
		<tr>
			<td>Residential Delivery:</td>
			<td>
			<input type="checkbox" name="ResidentialDelivery" value="1" class="clearBorder" <%=#@~^KwAAAA==210m;tm3}wDrW	`EI/r[xYbCVG+Vb-nDHJBPEFEb/Q4AAA==^#~@%>>
			</td>
		</tr>
		<tr>
			<td>Payor:</td>
			<td>
				<select name="PayorType" id="PayorType">
				<option value="SENDER" <%=#@~^JgAAAA==210mU+^+mDraOkKxcJhlzGMKXanJBJ?AH92"J*5AwAAA==^#~@%>>Sender</option>
				<option value="RECIPIENT" <%=#@~^KQAAAA==210mU+^+mDraOkKxcJhlzGMKXanJBJIA/(n&2gKE#xg0AAA==^#~@%>>Recipient</option>
				<option value="THIRDPARTY" <%=#@~^KgAAAA==210mU+^+mDraOkKxcJhlzGMKXanJBJK_(]fhb"KeJbLg4AAA==^#~@%>>3rd Party</option>
				<option value="COLLECT" <%=#@~^JwAAAA==210mU+^+mDraOkKxcJhlzGMKXanJBJZ}JJ2;Kr#KQ0AAA==^#~@%>>Collect</option>
				</select>
				<%#@~^JwAAAA==21/m"+$EkM+9(:monKmo~EhlXK.KHw+rS~0mVk+TQ4AAA==^#~@%>
			</td>
		</tr>
		<tr>
			<td>Payor Account Number:</td>
			<td>
				<input name="PayorAccountNumber" type="text" id="PayorAccountNumber" value="<%=#@~^LgAAAA==210mwk^VsKDsokV[`rnCzKDb1^W!xYg;h4Dr~~0Csk+bxBAAAA==^#~@%>">
				<%#@~^MAAAAA==21/m"+$EkM+9(:monKmo~EhlXK.b1mW!UO1!:(+.JS~6ls/4REAAA==^#~@%>		
			</td>
			</tr>
			<tr>
			<td>Payor Country Code:</td>
			<td>
				<input name="PayorCountryCode" type="text" id="PayorCountryCode" value="<%=#@~^LAAAAA==210mwk^VsKDsokV[`rnCzKDZK;xDDX;G[+r~,0CVdn*/Q8AAA==^#~@%>">
				<%#@~^LgAAAA==21/m"+$EkM+9(:monKmo~EhlXK.ZKExD.zZKNJSPWC^/nGhEAAA==^#~@%>		
			</td>
		</tr>			
		<tr>
			<td></td>
			<td>
			<input type=submit name="submit" value="Request Email Return Label" class="ibtnGrey">
			</td>
		</tr>
	</form>
</table>
<%#@~^QwAAAA==@#@&B&JPG2?:I}eP:C3Pw293oPr~B2;K@#@&dnY,W(Lo+[3XZslkdP{PUGDtkUL@#@&VREAAA==^#~@%>
<!--#include file="AdminFooter.asp"-->