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
<!--#include file="../includes/dateinc.asp"-->
<!--#include file="../includes/currencyformatinc.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="AdminHeader.asp"-->
<%#@~^+QoAAA==~@#@&/KxkYPbnmL+Uk"+{*@#@&@#@&fbhPbnlTn/EMDxO~~^KxUYhwBP.dBP\C.wVCo&U1WhaVY+BP!nDH~~/DD6]G~Pa^\|kxD6.ND&f@#@&@#@&@#@&@#@&E=U?U==?UU==?U=U?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U=@#@&B~?:)I:)~6gPS6)G@#@&B?=?U=?U?UU?U?=U?U=U?U==?UU?=U?UU?==U?U?U=U==?U=U?=U?U==?UU@#@&@#@&vzJ~U2P,nzM2,1z\2U@#@&w1nCL1lsnP{PJwn[2X{tlUlLnUtrwsnxD/]nkEVOdcldwr@#@&2.Mnmo+glsnP{PEsN3a|Hl	Co?tb2h+	YkK.l^Vcldwr@#@&@#@&v&JPrK3gP9b:)~bjA@#@&ml^V,Gwx94v#@#@&@#@&BJ&PU2K,Pu2,sAf3(~6~93Z:@#@&k+O~K4Lon92aZ^Ck/~{Pg+h,w1o+92aZ^ldd@#@&@#@&vzJPMAP~nzMAPHj\$AI@#@&bWPM+5;/Y 5!+.XkOMkUT`rknmo/EMDnxDJbxrJPD4+	@#@&irKlT+;E.DnUD'qP@#@&+^/n@#@&dkKCT+/EM.xO{I;E/D p!+.XUY.r	o`rrnmo+;;.DxDJb@#@&n	N~k6@#@&@#@&v&JP?6]:P6IG3"@#@&kYMrIG'Mn;!+dYvJG.9+Drb@#@&k0,dOD}IG'EJ~O4+U@#@&7/DD6]G'J2^hl^3mLqU6W|?tbwanNGlO+,f3j;~Pb[rMN+ME@#@&Ax9P(0@#@&kY.?K.Y{Dn5!+/Ocr/GDDE*@#@&b0,/YM?K.Y{JEP:tnU@#@&dkODUWDDxEfA?;J@#@&3U9P(0,@#@&@#@&v&JPM3P,r]fA],q9@#@&am\|/D.rMNnD&fx];EdYvJk9Eb@#@&w1\m/O.U+d/bGx}D[nMqfxj/dkKUvJ21b9:k	rM[+Mq9J*@#@&r6Pw1-{kYDUnd/bW	r.Nn.&fxJr~r"Psn	`w^-|/OD}.9+.&f*@*!,Y4nx@#@&7w1\mr	YrM[+Mqf{2^\|/DD6D[nMq9@#@&7?/drKx`E21b[:bU}D[D&fJ*'a^\|kUY}D[nMqf@#@&+^/+@#@&7w1\|kUY6.9+.qGxw1\mdDD?ndkkGx}.9+.&f@#@&+	N,r0@#@&vD/2G	/+cADbY+,jn/kkKxcJ2^zNhk	6D9+.(GJ#@#@&EDn/aG	/nc+	N@#@&@#@&vPV2PP:C3~hbZn)MA?@#@&v~@*@*@*,KC4snk)~w1Kl13CLqxWG@#@&5E.HPx,diJ?ASA/K,w^nmmVCT+q	WWcePr@#@&;!+MX~'~5!+.X,'Prs]6tPw^KmmVlTn&xWKPr@#@&$E.X,'~;!+.z,[Pr	CAI2,2^nmm0lL+(U6W k96D9+.xrP[~21\mk	O}D[D&fPLJ,Ed@#@&@#@&EP@*@*@*PZKUNbYkKUd@#@&q6P]+5;/OR5;+MXjOMkxLcrKzwjl.1tr#'rk96D9+.J,K4n	@#@&iO+sw;MzjYM'"+5EndDR}E.XUY.r	o`EC9\5E.HJb@#@&ik0,Yhw$Dz?DDxErPY4nx@#@&diOn:a;MXjY.xrP6IG3I,Ae~r[PdOMr]f,'rPELPkYDUWMO@#@&dnVk+@#@&idYhw$DXUO.'vk	YcYnha;.XUOD*PR~kmw.n*@#@&di5!+.H'$E+MX,'PrP	CAI3~bNrM[+MPS&F3PE]rP'Pm@#@&d7Yhw$DzjDDP'~r]vP}]G2],AIPJLPkOD}I9PLJ~ELP/D.?KDY@#@&7+	N,kW@#@&3	N~q67@#@&B(W,I+5;/OR5;DzUYMkxT`rPXa+j+mD^4r#'rGD9+DkOCY!/rPPtnU@#@&vd$;+MXx5!+Dz~LPEPquAI3,WMN+M/crN;EdYK:n.{mEkOWs+Dk rN;EkYG:n.,bHf,GD9+.dDlY;d,S(|A~E]E,[,{@#@&Bi]+$En/DR};DXUODbxovECN7;!+.XEb,[~JuvP}I93"PAe~r[~/D.}I9,[rPJLPkODUW.Y@#@&v3	NP&W@#@&@#@&knOPM/{?nD-nMR/DCYr8%mYcEzf6f~ "+^KD9/+DJ*~@#@&@#@&DkR/;M/WMJW1lYbGU'mN`/nZsrxO@#@&./cZC^4+?r"'rnmL?r.+@#@&DkRhCo?ry'rKmo+Ury@#@&Md ra+	P5En.H~~mKUxD+h2@#@&@#@&r6PnDM 	Eh(+MP@!@*PZ~Y4+U@#@&d^C^VPMdR;VWkn@#@&i/Y~Ddx	WOtbUo@#@&7./wGUk+ D[bDn1Y,JYm43DMRC/agn.MWD{E[,?+M-nDcjMVnx^G9+cJA.DKD~r	PVrdDW.N.k)~r[ADDcfdmMk2YbWUb,@#@&UN,q0@#@&@#@&M/cHG\nobDdY@#@&@#@&B&&,M2P~tbpPh)V2j@#@&Gk:,khCoZGE	Y@#@&bnlTnZKExDx./cnmonZG;	Y@#@&&WP;kUOvknCLZ;DMn	Yb,@*,Zk	YvrnmonZKEUO*PK4nx,knmLnZ!DM+UYxrhlL+;GE	Y@#@&&0PrKmonZ!.M+UDP@!PF,K4nx,kKlT+/;MD+	O'8@#@&@#@&vzJPU2PP)$UrJj:3Phb!3@#@&Dd z4dW^;D+Kmo'khlTnZ!D.+	Y@#@&@#@&BJ&PGq?hJ)5,2"I6I~\UM@#@&sdo{Dn5!+/O $EnDHdDDr	ovJ:korb@#@&@#@&k6PhdT@!@*rEPDt+	~@#@&iIkQDAA==^#~@%>
	<div class="pcCPmessage">
		<img src="images/pcadmin_note.gif" alt="Alert" width="20" height="20"> <%=#@~^AwAAAA==hkoRwEAAA==^#~@%>
	</div>
	<%#@~^WwAAAA==~@#@&n	N,k0,@#@&@#@&@#@&vzJP9(UnSzeP_2bG3]@#@&k6P./ nK0~Y4nx,@#@&7aD+d;^Yd'rTr@#@&Vk+P@#@&0hIAAA==^#~@%>
	<table class="pcCPcontent">
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<th colspan="3">Manage FedEx<sup>&reg;</sup> Shipments for Order Number <%=#@~^GwAAAA==ckm2M+3kxD`a^\|kUY}D[nMqf*bsQkAAA==^#~@%><a name="top"></a></th>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr> 
			<td width="25%" align="left" valign="bottom"> 
			<%#@~^FQEAAA==~@#@&7idiBzJPU4WSkUo,YGOmVP	;:(+D,GWPalT+dPWG!x[PmUN,Y4n,mE..xOPaCT+~	Es4+M@#@&7did]+kwGUk+Rq.kD+Pr9r/aVmXrxL~hlL+,@!4@*J~',knCLZ;DMn	Y~LPr@!z(@*,G0,@!8@*rP'~bnlTnZKExD~'Pr@!J4@*@!8.@*J@#@&i7diIndaWxdnc.kDn,JPKYmVPUtb2:xO/,sG;	N),@!4@*JPL~./cImGD[/KEUY,'Pr@!&8@*JP@#@&id7dtEwAAA==^#~@%></td>
			<td width="50%" align="center">
				<img src="images/fedexCompatible.gif">
			</td>
		    <td width="25%" align="right" valign="bottom"><input type="button" name="Button2" value="Closeout & Print Manifest" onClick="document.location.href='FedEx_ManageShipmentsClose.asp?PackageInfo_ID=<%=#@~^DgAAAA==21\mbxDrD9+M(ffAUAAA==^#~@%>'" class="ibtnGrey"></td>
		</tr>
	</table>
<%#@~^CAAAAA==~x[,k6PZgIAAA==^#~@%>

<table class="pcCPcontent">

	<form name="checkboxform" action="FedEx_ManageShipmentsTrack.asp?id=<%=#@~^DgAAAA==21\mbxDrD9+M(ffAUAAA==^#~@%>&action=batch" method="post" class="pcForms">
		<tr> 
			<th nowrap>Shipped</th>
			<th nowrap>Tracking</th>
			<th nowrap>Contents Description</th>
			<th nowrap>Package Details</th>
			<th nowrap>Options</th>
			<th nowrap>Returns</th>
			<th nowrap>Select</th>
		</tr>
		<%#@~^LgAAAA==~Gkh,:1WE	Y@#@&di:^W!xOxZ@#@&i7q6PDk 3rwP:tnx~0AsAAA==^#~@%>
			<tr>
			<td colspan="11">
				<div class="pcCPmessage"><img src="images/pcadmin_note.gif" width="20" height="20"> No Results Found</div>			</td>
			</tr>
		<%#@~^wgMAAA==~AVd@#@&ddiB,jtKhrxTP.n^+\mUY,D+1G.Nk@#@&d7d9rsPdYM/W^@#@&7id/O.;Ws'r:AF3828JP@#@&i7dGkhPMmG;	Y~,r~,6@#@&77d@#@&id7sG.,kxF,PW,Dd hlonjbyn@#@&7id7@#@&iddiD1GE	Yxk@#@&77idq6~m!DDUOnmoP@*Pq~:tnx@#@&did77wWD~a{F~KK~vm;MDxYhlTnP Pq#@#@&77iddi.mKExDxq!,_,D^W;UD@#@&di7di1naD@#@&77id3x9~&0@#@&P,PP,P,~P,P~P,P~~,PP,~P,PP@#@&7did&0~1GO,DdRA6s,K4n	P@#@&7id7d&W,/OMZKVP@!@*,EawsoswsE~:t+	@#@&iddi77/DD;Ws'E:wsoswoJ@#@&77idd3sk+~@#@&7id7idkYD;W^xJ[2q282qE@#@&di7di2x9~(0@#@&id7d7@#@&d7di7w1\mr	YnC^0lL+&U6Wm&f{D/vJa^nmmVlT+(U6W{&9J*@#@&i77diw1\mkUO}D[+MxDk`Er9rD[nMJb@#@&7id7iwbNnmm0Co1;:(+.xM/`r2mhlm0CL+&x6WmnC^0lL+g;:(+.E*@#@&77id7w1-|/OMKMlm0k	L1!:8+M'.dvJw1Kl13lTn(x6W|K.l^VbxL1!h4DEb@#@&d77id2k9KmmVmo+bo4O'M/cJamKC13lTnq	0W|KCm0lT+	+rL4YE#@#@&did77akNKC13Coj4k2a+9flD+{./vJ2mhl^Vmo+&U0K{?4r2wNGlO+Ebid7di@#@&id77iwm-mkY.?4raHnDtKNP{PMd`rw^nmmVCT+q	WW|?tb2\+DtKNE#@#@&id7di2m7{dOMsfp]mYn'MdvJ21nmm3mo(x6WmsG(]CD+J*@#@&iddi7@#@&idid7:^G!xO's^W!xOQ8P@#@&7id7dqfEAAA==^#~@%>
												
					<tr> 

						<td bgcolor="<%=#@~^CAAAAA==~kY.;W^PtwIAAA==^#~@%>"><%=#@~^IwAAAA==j4WAGlD+sM:DcwbNKl13CL?tb2wNfmOn#PQ0AAA==^#~@%></td>
						<td bgcolor="<%=#@~^CAAAAA==~kY.;W^PtwIAAA==^#~@%>">
							<a href="FedEx_ManageShipmentsTrack.asp?id=<%=#@~^DAAAAA==21\mbxDrD9+M7wQAAA==^#~@%>&PackageInfo_ID=<%=#@~^FQAAAA==21\mbxDnl13mL+&xWW|q9FwgAAA==^#~@%>"><%=#@~^FQAAAA==21\mkYMKDmm0rxT1;:(+.nQgAAA==^#~@%></a>						</td>
						<td bgcolor="<%=#@~^CAAAAA==~kY.;W^PtwIAAA==^#~@%>">
						<%#@~^KQMAAA==@#@&d7ididB,MAPP:C3Phb/FzM2,/rgK2gPj@#@&did7d7v,@*@*@*,Pl(Vnd=Pw.G9E^YkS,n.KN!mYkrM[+M+[@#@&d77idd$;+MXP{~7dr?AS3ZP~hDGN!^Ykr.[D+[ amKl1Vmon&x6W{&f,SPaDGN!mOdcN+k^DbwYbGU~,wMW[E^OkRrNh.W9E^O,PJ@#@&id7di7$EnMX,'P$E.X,[~JwI6\,nDK[E1Y/}.[+M+9PE@#@&7id7di5EDz~{P;;nMX~[,E&1HAI,9r&1,2DKN;mD/~E@#@&di7did;!n.X,',;;+.z,[~J}HPhDG[!mYd6MNnD[ck[hDKNE1Y,xPaDGN!mOdckNh.W9EmD~E@#@&did7d75!+.X,xP$En.HP[~EqC3IA~hDG9E1Y/}D9nDN w1nC^0lo(x6W{&9xJ,[,w^\mr	YKl1VlT+(U6W{(9,[EPr@#@&d7ididdidi7d@#@&7did77k+Y,./y'/.-+MR;DnlOn}4%+1O`rb96GAR]n1W.NUnDJb@#@&iddidid+DP./y'^G	xYhwc+6^;Y`$EnDzbid@#@&i7did7@#@&dd77idr0,nMD 	Es4+M@!@*TPDtnx@#@&77iddi7BJzP4CUN^+,l[:rU,+.DK.@#@&d77iddnU9Pr0@#@&id7idi@#@&idi7dikWPgrP~M/ cnW6PY4nU@#@&did7d77GW~E	Ok^P.dyR+GWi@#@&di7id7idam\|/D.nMW[E1Y9nkmDb2YbWx,x~Dk vJ[+d^Mk2YbGxr#@#@&idd77id7btQAAA==^#~@%>
								<%=#@~^GQAAAA==21\mkYMnDKN!^YG+dmMk2ObWxZgoAAA==^#~@%><br />
								<%#@~^RgAAAA==@#@&d7ididdM/y :K\nx6O@#@&ddi7didSKG2@#@&did7d7n	N~k67did77@#@&d77id79wkAAA==^#~@%>						</td>
						<td bgcolor="<%=#@~^CAAAAA==~kY.;W^PtwIAAA==^#~@%>">
							
						
							<table width="100%" border="0" cellspacing="0" cellpadding="0">
							  <tr>
								<td align="right">Weight:</td>
								<td align="left"><b><%=#@~^EAAAAA==2bNKmm0loro4YUQYAAA==^#~@%> lbs.</b></td>
							  </tr>
							  <tr>
								<td align="right">Method:</td>
								<td align="left">
									<b>
										<%#@~^tQAAAA==@#@&d7ididdidir0,w^\|/O.Utka\+DtW9~xPrJ,Y4+U@#@&d7di7did77iD+d2Kxd+cAMkOPrMDKE	[J@#@&7did77iddinVk+@#@&77didid7d77M+dwKU/RA.bY+~21\m/D.UtraHYtKN@#@&did7did77i+x9~k6@#@&i77didid7ddSYAAA==^#~@%>
									</b>
								</td>
							  </tr>
							  <tr>
								<td align="right">Net Rate:</td>
								<td align="left">
								<b>
								<%#@~^vQAAAA==@#@&d7ididdik6~w1\m/DDo9oIlDnP@*P!,O4+	@#@&d7d77id7dMn/aWUdRh.rD+~/1/!Djbo	[:Kxz`am-{kY.oG(ImO+*@#@&i77dididnVdn@#@&7di7did77M+/2G	/nRS.bYn,JzVYD	CYPKlHW.E@#@&di7diddinUN,k6@#@&d77id7didCwAAA==^#~@%>
								</b></td>
							  </tr>
						  </table>						</td>
						<td bgcolor="<%=#@~^CAAAAA==~kY.;W^PtwIAAA==^#~@%>">
							<a href="FedEx_ManageShipmentsPrinting.asp?path=FedExLabels/Label<%=#@~^FQAAAA==21\mkYMKDmm0rxT1;:(+.nQgAAA==^#~@%>.PNG" target="_blank">View/ Print Label</a><br />
							<%#@~^QQAAAA==~b0~K4Ns+92X/Vm/dRamWmw+NAa?hrfv2^\|k	YKl^Vmonq	WW|q9b{YD;n,Y4+	~KBcAAA==^#~@%>
								<a href="FedExLabels/SPOD<%=#@~^FQAAAA==21\mkYMKDmm0rxT1;:(+.nQgAAA==^#~@%>.PDF" target="_blank">View SPOD</a>
							<%#@~^BgAAAA==~VdP6QEAAA==^#~@%>
								<a href="FedEx_ManageShipmentsSPOD.asp?id=<%=#@~^DAAAAA==21\mbxDrD9+M7wQAAA==^#~@%>&PackageInfo_ID=<%=#@~^FQAAAA==21\mbxDnl13mL+&xWW|q9FwgAAA==^#~@%>">Signature POD</a>
							<%#@~^CAAAAA==~x[,k6PZgIAAA==^#~@%>						</td>
						<td bgcolor="<%=#@~^CAAAAA==~kY.;W^PtwIAAA==^#~@%>">
							<a href="FedEx_ManageShipmentsCancel.asp?id=<%=#@~^DAAAAA==21\mbxDrD9+M7wQAAA==^#~@%>&PackageInfo_ID=<%=#@~^FQAAAA==21\mbxDnl13mL+&xWW|q9FwgAAA==^#~@%>">Cancel Shipment</a><br />
							<a href="FedEx_ManageShipmentsEmail.asp?id=<%=#@~^DAAAAA==21\mbxDrD9+M7wQAAA==^#~@%>&PackageInfo_ID=<%=#@~^FQAAAA==21\mbxDnl13mL+&xWW|q9FwgAAA==^#~@%>">Email Return Label</a><br />						</td>
						<td bgcolor="<%=#@~^CAAAAA==~kY.;W^PtwIAAA==^#~@%>"><input type=checkbox name="check<%=#@~^BgAAAA==h1W;	YlgIAAA==^#~@%>" value="<%=#@~^FQAAAA==21\mbxDnl13mL+&xWW|q9FwgAAA==^#~@%>"></td>
					</tr>
					<%#@~^JQAAAA==~M/ tW7+16D@#@&id7dAx[~&0@#@&7di1+XOWwgAAA==^#~@%>
			<input type=hidden name="count" value="<%=#@~^BgAAAA==h1W;	YlgIAAA==^#~@%>">								
		<%#@~^CAAAAA==~Ax[,q6PJgIAAA==^#~@%>
		<tr> 
			<td colspan="11">
				<%#@~^EQAAAA==r6Ph1W!xY@*!,OtxwgUAAA==^#~@%>
					<a href="javascript:checkAll();"><b>Check All</b></a><b>&nbsp;|&nbsp;<a href="javascript:uncheckAll();">Uncheck All</a></b><br>
					<br><input type=submit name="submit" value="Track All Selected Packages" class="ibtnGrey">
					<script language="JavaScript">
					<!--
					function checkAll() {
					for (var j = 1; j <= <%=#@~^BgAAAA==h1W;	YlgIAAA==^#~@%>; j++) {
					box = eval("document.checkboxform.check" + j); 
					if (box.checked == false) box.checked = true;
						 }
					}
						
					function uncheckAll() {
					for (var j = 1; j <= <%=#@~^BgAAAA==h1W;	YlgIAAA==^#~@%>; j++) {
					box = eval("document.checkboxform.check" + j); 
					if (box.checked == true) box.checked = false;
						 }
					}
					
					//-->
					</script>
				<%#@~^BgAAAA==n	N~b0JgIAAA==^#~@%>			</td>
		</tr>
	</form>
              
	<tr>
		<td colspan="11"> 
			<%#@~^GQAAAA==~b0~aI/E^Yk@!@*r!EP:tnU,LgcAAA==^#~@%>
				<table width="100%" border="0" cellspacing="0" cellpadding="4">
					<tr> 
						<td> 
							<form method="post" action="" name="" class="pcForms">
							<b> 
							<%#@~^aAAAAA==~"+daW	/+cMrY`E@!6WUO,/k.n'yP0m^n'mDbls@*KCT+~JL~khlLn;ED.n	Y~[,E,WW,JLPkhlTnZKEUY,[~E@!z0KUY@*@!n@*EbIR8AAA==^#~@%>
							<%#@~^gAAAAA==~Efrkw^lX,1aY,z~nM+-~(EYDGxk@#@&i77didikWPrKmonZ!.DxO~@*PF~O4+U@#@&7id7ididBq+,CDPUWDPCO,Yt~4ok	UrxT~,/4WA~DtnPa.+7P8;DYWU~/CQAAA==^#~@%>
								<a href="FedEx_ManageShipmentsResults.asp?id=<%=#@~^DAAAAA==21\mbxDrD9+M7wQAAA==^#~@%>&TypeSearch=<%=#@~^IQAAAA==.;;/DR;!+Mz/DDrxT`EPHw+UnlMmtrbkQwAAA==^#~@%>&advquery=<%=#@~^HwAAAA==.;;/DR;!+Mz/DDrxT`EC9\;!nDHJ#CgwAAA==^#~@%>&FromDate=<%=#@~^IQAAAA==.;;/DR;!+Mz/DDrxTPcEwDWs9lD+J*~6wsAAA==^#~@%>&ToDate=<%=#@~^HwAAAA==.;;/DR;!+Mz/DDrxTPcE:WfmO+r#PGgsAAA==^#~@%>&iPageCurrent=<%=#@~^DgAAAA==rhlLZ!DDxDRFJwUAAA==^#~@%>&order=<%=#@~^BgAAAA==dDD6"fPgIAAA==^#~@%>&sort=<%=#@~^BwAAAA==dDDjKDDAQMAAA==^#~@%>"><img src="../pc/images/prev.gif" border="0" width="10" height="10"></a> 
							<%#@~^cAAAAA==~x[,q6@#@&idi7did(0,kKCT+ZK;xDP@!@*~qPDtx@#@&77id7di7sKD~({FPPG,kKlTn;W;	Y@#@&didi7did7q6P(xbnlTnZ!DDUOP:tx~mBkAAA==^#~@%>
										<%=#@~^AQAAAA==(SQAAAA==^#~@%> 
									<%#@~^BgAAAA==~AVdPyQEAAA==^#~@%>
										<a href="FedEx_ManageShipmentsResults.asp?id=<%=#@~^DAAAAA==21\mbxDrD9+M7wQAAA==^#~@%>&TypeSearch=<%=#@~^IQAAAA==.;;/DR;!+Mz/DDrxT`EPHw+UnlMmtrbkQwAAA==^#~@%>&advquery=<%=#@~^HwAAAA==.;;/DR;!+Mz/DDrxT`EC9\;!nDHJ#CgwAAA==^#~@%>&FromDate=<%=#@~^IQAAAA==.;;/DR;!+Mz/DDrxTPcEwDWs9lD+J*~6wsAAA==^#~@%>&ToDate=<%=#@~^HwAAAA==.;;/DR;!+Mz/DDrxTPcE:WfmO+r#PGgsAAA==^#~@%>&iPageCurrent=<%=#@~^AQAAAA==(SQAAAA==^#~@%>&order=<%=#@~^BgAAAA==dDD6"fPgIAAA==^#~@%>&sort=<%=#@~^BwAAAA==dDDjKDDAQMAAA==^#~@%>"><%=#@~^AQAAAA==(SQAAAA==^#~@%></a> 
									<%#@~^CAAAAA==~Ax[,q6PJgIAAA==^#~@%>
								<%#@~^BgAAAA==~g+aDP3wEAAA==^#~@%>
							<%#@~^CAAAAA==~x[,k6PZgIAAA==^#~@%>
							<%#@~^ZgAAAA==~b0~;q	Y`bnmL+;E.DxOb,@!@*,/q	Y`bKCoZKEUYb~Dtnx@#@&did77iddv	PCD~	WO,lDPY4+,nx9~~/4WA~mPxaY,Vk	V~4R0AAA==^#~@%>
								<a href="FedEx_ManageShipmentsResults.asp?id=<%=#@~^DAAAAA==21\mbxDrD9+M7wQAAA==^#~@%>&TypeSearch=<%=#@~^IQAAAA==.;;/DR;!+Mz/DDrxT`EPHw+UnlMmtrbkQwAAA==^#~@%>&advquery=<%=#@~^HwAAAA==.;;/DR;!+Mz/DDrxT`EC9\;!nDHJ#CgwAAA==^#~@%>&FromDate=<%=#@~^IQAAAA==.;;/DR;!+Mz/DDrxTPcEwDWs9lD+J*~6wsAAA==^#~@%>&ToDate=<%=#@~^HwAAAA==.;;/DR;!+Mz/DDrxTPcE:WfmO+r#PGgsAAA==^#~@%>&iPageCurrent=<%=#@~^DgAAAA==rhlLZ!DDxDQFJQUAAA==^#~@%>&order=<%=#@~^BgAAAA==dDD6"fPgIAAA==^#~@%>&sort=<%=#@~^BwAAAA==dDDjKDDAQMAAA==^#~@%>"><img src="../pc/images/next.gif" border="0" width="10" height="10"></a> 
							<%#@~^IgAAAA==~x[,q6P@#@&di7did7mmVs~1VWknf(`#,hQcAAA==^#~@%>
							</b> 
							</form>						</td>
					</tr>
				</table>			</td>
		</tr>
		<tr>
			<td colspan="11" align="center">
			<%#@~^SAAAAA==~@#@&7idam\|/D.nM+-kKEdKmo+,xPrrD9[nYmk^/ ld2Qk['r~[,w^-|kxO6MNnD@#@&id7OxUAAA==^#~@%>
				<input type="button" name="Button" value="Go Back To Order Details" onClick="document.location.href='<%=#@~^EwAAAA==21\mkYMnD\bGEknCo2wcAAA==^#~@%>'" class="ibtnGrey">
			<%#@~^VQAAAA==~@#@&7idV/P@#@&did2m7{dOMnD-kKE/hCL+,',J6D[[YCk^dRm/2_bN'E~LP]+$;/OvJbNJ*@#@&7dikhYAAA==^#~@%>
				<input type="button" name="Button" value="There are no Packages to Display >>> Go Back" onClick="document.location.href='<%=#@~^EwAAAA==21\mkYMnD\bGEknCo2wcAAA==^#~@%>'" class="ibtnGrey">
			<%#@~^CAAAAA==~x[,k6PZgIAAA==^#~@%>		
			</td>
	</tr>
		<tr>
		  <td colspan="11" align="center"><br /><%=#@~^IAAAAA==~amW|sN2XMrYSnomV9rkmVmr:D/,BAwAAA==^#~@%></td>
		  </tr>
</table>
<%#@~^QwAAAA==@#@&B&JPG2?:I}eP:C3Pw293oPr~B2;K@#@&dnY,W(Lo+[3XZslkdP{PUGDtkUL@#@&VREAAA==^#~@%>
<!--#include file="AdminFooter.asp"-->