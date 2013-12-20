<%@ LANGUAGE = VBScript.Encode %>
<%#@~^vwEAAA==@#@&BP4kkP0bV~kkP2lMY~G6PnMGN!mY;C.YBPmx~+^Gs:nD1nPmw2sbmlOrKx~N-VGa+9Pl	N,dW^N~4HP3CMVX,(:almD~JS;R,n.W[;1Y/lMO~,kOd,/W;.1+~mK[~~DtPnMW9;mDZCDDPUCs+PmUN,VWTG~lM+,w.W2nMYzPKWPAl.sHPqh2mmO~,JdZ ,ZKwXMkT4Y, T!8O+TZ&R,)V^PDbL4YkPM+d+.-N PIGE,l.n,xWO~mVsWSn9POKP!/+BPmsYDSP9kdOMk4!O+,lx9&GD,D/nVs~mxzPaCDD/~G6Pn.G9E^Y;CMYvkPkWEMm~mKNnPSkO4KEY,OtPhMrOYx,mGxdn	Y~W6~2mDsz,q:2C1Y P:G,mG	YmmY,2m.VHP(:al^OBPw^nlk+P7rdkDPShARnCMVzks2l1Y ^K:R@#@&MpwAAA==^#~@%>
<%#@~^JwAAAA==~alLKbYV'rjtbw2k	o~/xY.P6WD,onNA6rPdg0AAA==^#~@%>
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
<%#@~^/xYAAA==~@#@&9b:,;EDHSPM/SP1WUUD+:a@#@&Gk:,rKlT+;E.DnUD~~\m.s^lL(	mWh2^+O+B~!+.H~,/YMr"9~,w^\|kUO}DN.qG@#@&GrhPam7{dY.\Y4W9Hls+S~am\mdDD\+D4KN]w^X~,Z!dYK:nD:DCUklmDrW	qNUOk6kDSP2^7{dYM)m1W;UD1Eh8DSPa^7{dDDt+YDg;:(+.~,w^-|/YM/lMDk./W9+@#@&9kh~am-{kOD:DC^0kxLH!:8+MS,w^7{kYDUtb2:xOb1mG;	Y1!h4D@#@&9r:,w1\m/O.G+dYbUlDkGU;WEUOMX/W9nBP21\|/YMfdYbxCYbWUKK/YmsZKN+B~2m7{kY.SCUTECo/W9+S~am\mdDDJW1C^+/KN~Pam7m/DD9+DlrsUml	d~,wm7mdYMnmorxLPK3nx@#@&fb:~WN+amaWdY9CDlS,W(LsNAaZ^ld/BPG8NrED2ED(Hd9GmBPkD-s39A(p:^uYDwS~w2f3p|Dn/!sD~~w2G2(|j"J~,w^\|/O.ADDK.Hko~,2^\|/DD)mOrKx@#@&@#@&@#@&@#@&v?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?@#@&EPjPzIKl~}1~S})G@#@&EU?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U@#@&@#@&vJz,nbV2,Hbt2@#@&amKCT+1mh+{Js[36|HmxConj4k2:UYkZC2DEDn m/2J@#@&AD.hlT+1m:xJw+[2X{\C	lojtbw:UO/;laY;Dn m/2J@#@&@#@&B&&,bZP(}1@#@&a^7{dDDzmYbW	~',Dn;!+dOvJb1OkKxJ*@#@&@#@&BJz~rK3gP9b:)Az?3@#@&mlss,W2+	9(`b@#@&@#@&BJz,j2:PPCAPo3G2(,6Ax2Z:@#@&/Y,W8Lon92aZ^C/kPx~g+h~21snNAa;VCk/@#@&d@#@&E&z,s3fA(~/"2fAHK&bSU@#@&;!+MX~'~7iJj2d3Z:Pj4bw:nUDKzwdcEdD&f~,?4rws+UY:X2nkRwmd/SWD9S~?4ka:nxOPHwn/c)m1+dddkmnUk+~J@#@&$EnMX,'P$E.X,[~JwI6\,?tb2:xY:z2+kPr@#@&;;nMX~',5EDz~LPJ	uAI3Pvcv?4bws+xDKH2+kRrNUtr2s+xDb'8##pE7@#@&/Y~Ddxk+.\.R;DnCD+r8%mO`r)Gr9~R"+mKD9j+DJb@#@&/nO,D/{^W	xYh2R6m;Ync$EnDHb@#@&kW~grK~.kRnW6~Dtn	@#@&dw1\|dYMb^mKEUOgE:(nD{D/vE;/D&fE#@#@&iw^\|dYMHnOD1;h(+.'MdvJ2m/khWMNrb@#@&d2m7{dOM2x7rDKx:UOP{PrK3?PE@#@&nx9~k6@#@&dYP.d{xGY4r	o@#@&@#@&BzJPU3K,I3p`q]3GP.z]qzASAj@#@&am7{dY.\Y4W9Hls+~x,Js9pU?K..kkG	ZmwY!D]+$En/DJ@#@&am\|dYMH+D4GN"+aVzPx~rs9(Ujnj+.dbWx/CaY;D]wsHJ@#@&@#@&@#@&vU?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?@#@&EPAHf=P61,S6)G@#@&E=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U@#@&@#@&@#@&veCeeCMMeCeCeMeMMCeMeCMeCeMMCeeMMCeMeCMCeMCeCeeCeCMeCeMeCeMMCeeCMeCeeCMMeCeC@#@&B~j:b]K=~n}?P~~bZF@#@&BMeCMCeMCeCeeCeCMeCeMeCeMMCeeCMeCeeCMMeCeCeMeMMCeMeCMeCeMMCeeMMCeMeCMCeMCeCe@#@&k6~D;;+kY WKD:vE/!4:bOE#@!@*rJ~Y4n	@#@&@#@&7B?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U@#@&iv,MnDPmVV,W6~Y4+~D;;rM+N,rx6WDsCOkKxc@#@&dv=?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U=7@#@&d@#@&iB&z,!xnMk1P+MDK.P6W.PalLn@#@&da^\|/YM!nxDbmKlLnAD.WM~',J)O,V+CdDPGx~M+5!kM+N,0bnV9PAlkPnhaYXc~J@#@&d@#@&7B,/Dl.Y~jD-+MR?bNn~7lVr[mYrW	d@#@&7Ew1/{jl^rNmYnK6Oob+V97JzmmK;UYgEs4nDES,Y.ESPW!@#@&iBw^d|.CVb[mYn:+XYsb+^[drHnYDH;s4+ME~,YD!nSP8 @#@&7B2^k{#l^rNmYnP6YorV[dr/mD.b+MZW9+rSPDD;+BPqT@#@&d@#@&dEUU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=@#@&7EP/t^3,0G.,.lsr9lOkKU,2.MWM/R,fK~xKY~wMW^nNPbWPDt+Mn~lM+,+.DG.kR@#@&ivU?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?@#@&d&W,w^7{bxYADM@*!,K4+	@#@&7iD+k2W	/+c.nNbDmOP2^hlL+gC:P'~rg:dL{J~[,21\mkYMM+	+MrmhlL+AD.GM@#@&i3Vk+@#@&77d@#@&idvU==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U@#@&diB~$!kV[~}E.P:.mxdmmDkW	R@#@&diB=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?@#@&i7W(Ls[36;Vm/dRHnS(\S;CwDE.n,wm-mkY.HO4W[gls+~,w1-{kY.b1mG;	Y1!h4D~,2^\|/DD\+OnM1;:(nD@#@&77iW4%oN36;sm/dcMkY?bUo^+KlM+UO,JSK^lDkW	(9JBPrj!1)E@#@&7diG4Nsn[A6ZsCk/ MrD+jbxTV+hlMnxDPE.x[GMnDK[E1YqGESPr2&n/J77id7d@#@&didG8Ns+[3XZslkdc.bY?k	o^nnmDnxDPE#xNK.nMWN!^On^lD0GDhEBPEgpnUnT+r@#@&77iW8Lwn92a;Vm//cMrY?rxTVnKmD+	OPr.+	[GDhDKN;mO#DdkKUJBPET2!!E@#@&d7W(%w+[A6;Vlk/c3x9(\S:DCUklmDrW	Pw1-m/DDt+OtG[glh+i@#@&id@#@&idB&&,n.k	O,W;DPKED,xAVHPWWM:n[,D+$;+kYPXhs@#@&diB.+d2Kxd+cADbYn~6+Nna|wG/D[mYC@#@&idBM+k2W	/nRx[@#@&dd@#@&diBU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U@#@&7dv~U+UN,6EMPP.mx/C^DkGxc@#@&d7EU?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?@#@&id^l^V~G(Ls[2XZVmddRU+	NpHJ];;+kO`6+[nX{wGdDNCYmS,w^7{kYDAx7rDKxh+	Yb@#@&dd@#@&diBzJ~KDbxDPGEO~KE.PMn/aWUd@#@&77EDn/aG	/nchMkYPw3fA(mD/;sD@#@&i7BM+/aGU/Rx[@#@&7i@#@&divU?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?U@#@&7iB~SKC9P6!D,I+kwKU/R@#@&idv=?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U@#@&di^CV^PK4%sn[A6/Vmd/cSGC9(HJ]/;VDdvs3G2o{D/!sY*@#@&di@#@&7i@#@&i7B?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U@#@&idvP/4mVP6GD,+..KD/~WMWhPwn92ac@#@&ddEU?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=did7@#@&dd1CV^PW(%o+92XZslddc(\S"n/aWUd.+.r6Xc2M.hlL1m:+*@#@&7d@#@&7d@#@&77EzzJ&zJzzJ&&zJzJz&z&&Jz&zJ&zJz&&Jzz&&Jz&zJ&Jz&JzJzzJzJ&zJz&zJz@#@&idBJ&PdrMV(HM@#@&idvz&&Jz&zJ&zJz&&Jzz&&Jz&zJ&Jz&JzJzzJzJ&zJz&zJz&&JzzJ&zJzzJ&&zJz@#@&7dv&JPPDm^3bxL~gE:8nMPWWM~dWLk@#@&ddam7m/DDPDmmVr	o1!h4DP{~E.DkkGx/CaY;DmFr@#@&7iBz&~dWLPK;MPPMl	/l1YbGx@#@&7d1lss,W4No+926;sC/kRamd{JGTK.l	dl1YrG	`0n[6mwKdDNCDlBPw1\|dYMHnY4W[Hm:+LE{r[w1-m/DD:DCmVr	oHEs8+M[E bxJS~DD;+*@#@&d7EzJPSKo,GEMP]+kwGUk+@#@&7d1lV^~G4NsN36/sm/dRa^/|SGL:DlUdmmOkKUvs3G2o{D/!sYBP2m7{dOMH+D4W91lsn'J|JLw^\mdDDPDm^3bxLH!:4n.LJ W!Or~~DD!+#@#@&i7@#@&d7B?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU@#@&diBP"n[kM+1Y~hrO4PCPtn/klLn,rI~^K:2VOPdK:PYm/0 @#@&d7B?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU@#@&dik0,H6K,Vxcw^-|/ODA.DKD\dT#@*T~Dtnx@#@&id@#@&di@#@&idivU?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U@#@&id7v,?+D~r!DP"ndwKxk+~fCOmPOW,JW1ls @#@&d77EU=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU==?U=U@#@&id7EPz\lbVm8VP\+DtG[kPhbsV,/+m.^t,E	Vr:rON~V-+^/~G6P1G[/~4H~k+2mDmYk	o,UW9+dPSkO4,lPr&Jc@#@&i77B,F*P]+C["+dwKU/nC.xY@#@&id7B,+*P]l9I+kwKU/1GN@#@&7id@#@&7diBzJ~u2zfAI@#@&77iw^\|dYMIn2^XPx~K4%s[A6/^lk/R"+m[I/2W	/nHKN+vEzJI+aszCl9+.JS~rInw^zJ*d@#@&idd@#@&id7BJ&,2]"r"@#@&idi2m7{dYM2..KDZK[+,'PK8%sNA6/VCdkR]+m[I/2G	/+HG9+cJJ&AD.KDr~PrZK[+r#@#@&id721\{kODADDK.\+k/monPx~K4%s[2XZsCk/R]nmN]+k2Kxd1KN+vJJ&2MDGDr~~Et+/kCoJ#i77@#@&didGLkHAA==^#~@%>
			<table class="pcCPcontent">
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>
				<tr>
					<th colspan="2">FedEx<sup>&reg;</sup> Version Capture Complete</th>
				</tr>
				<tr>
					<td colspan="2" class="pcCPspacer"></td>
				</tr>		
				<tr> 
					<td colspan="2">
					<%#@~^XQAAAA==@#@&d7idik0,w1-{kY.Iwsz,'PrEPDt+	@#@&didid7DndaWU/ hMkOn,J?;^1+d/eE@#@&7idid+	N,r0@#@&7did78RQAAA==^#~@%>
					<%=#@~^DAAAAA==oAf3o{M+/!VDagQAAA==^#~@%>
					</td>
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
		<th colspan="2">FedEx<sup>&reg;</sup> Transaction Request </th>
	</tr>
	<tr>
		<td colspan="2" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
			<p>
			This completes the required version capture.
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
			<td></td>
			<td>
			<input type=submit name="submit" value="Capture Version" class="ibtnGrey">
			</td>
		</tr>
	</form>
</table>
<%#@~^QwAAAA==@#@&B&JPG2?:I}eP:C3Pw293oPr~B2;K@#@&dnY,W(Lo+[3XZslkdP{PUGDtkUL@#@&VREAAA==^#~@%>
<!--#include file="AdminFooter.asp"-->