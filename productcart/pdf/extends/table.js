var widths;
var cellaligns;
var cellborders;
var customborders;
var customstyle;
var cellheight;
this.cellheight=5;
this.Table = {Border:{Width:0.1,Color:''},Fill:{Color:''},TextAlign:"J"}
this.SetColumns=function(){this.widths=arguments;}
this.SetAligns=function(){this.cellaligns=arguments;}
this.SetBorders=function(){this.cellborders=arguments;}
this.SetCustomBorders=function(){this.customborders=arguments;}
this.SetCustomStyle=function(){this.customstyle=arguments;}
this.SetCellHeight=function(){var xdata=arguments; this.cellheight=xdata[0];}
this.HexToRGB=function(value){
var ar=new Array()
s= new String(value.toUpperCase());
ar["R"] = parseInt(s.substring(0,2),16);
ar["G"] = parseInt(s.substring(2,4),16);
ar["B"] = parseInt(s.substring(4,6),16);
return ar;
}
this.Row=function()
{
var xdata = arguments
var xi;var xh;var xnb;
var xw;
xnb=0;
for(xi=0;xi<xdata.length;xi++){xnb=Math.max(xnb,this.NbLines(this.widths[xi],xdata[xi]))};
xh=(xnb)*this.cellheight;
this.CheckPageBreak(xh);
for(xi=0;xi<xdata.length;xi++)
{
xw=this.widths[xi];
xx=this.GetX();
xy=this.GetY();
xy1=this.GetY();
if (this.Table.Border.Width>0||this.Table.Fill.Color!=''){
var xstyle='';
this.SetLineWidth(this.Table.Border.Width);
if(this.Table.Border.Color!=''){
var RGB = this.HexToRGB(this.Table.Border.Color);
this.SetDrawColor(RGB["R"],RGB["G"],RGB["B"]);
xstyle+='D';
}
if(this.Table.Fill.Color!=''){
var RGB = this.HexToRGB(this.Table.Fill.Color);
this.SetFillColor(RGB["R"],RGB["G"],RGB["B"]);
xstyle+="F"
}
this.Rect(xx,xy,xw,xh,xstyle);
var delstr=this.cellborders[xi]+"";
if (delstr.indexOf("B")>=0)
{
	this.SetDrawColor(255,255,255);
	if (delstr.indexOf("O")==-1)
	{
		this.Line (xx+this.Table.Border.Width,xy+xh,xx+xw-this.Table.Border.Width-this.Table.Border.Width,xy+xh) ;
	}
	else
	{
		if (delstr.indexOf("L")>=0) this.Line (xx,xy+xh,xx+xw-this.Table.Border.Width,xy+xh);
		if (delstr.indexOf("R")>=0) this.Line (xx+this.Table.Border.Width,xy+xh,xx+xw,xy+xh);
	}
}
if (delstr.indexOf("T")>=0)
{
	this.SetDrawColor(255,255,255);
	if (delstr.indexOf("O")==-1)
	{
		this.Line (xx+this.Table.Border.Width,xy,xx+xw-this.Table.Border.Width-this.Table.Border.Width,xy) ;
	}
	else
	{
		if (delstr.indexOf("L")>=0) this.Line (xx,xy,xx+xw-this.Table.Border.Width,xy) ;
		if (delstr.indexOf("R")>=0) this.Line (xx+this.Table.Border.Width,xy,xx+xw,xy) ;
	}
}
if (delstr.indexOf("L")>=0)
{
	this.SetDrawColor(255,255,255);
	this.Line (xx,xy+this.Table.Border.Width,xx,xy+xh-this.Table.Border.Width) ;
}
if (delstr.indexOf("R")>=0)
{
	this.SetDrawColor(255,255,255);
	this.Line (xx+xw,xy+this.Table.Border.Width,xx+xw,xy+xh-this.Table.Border.Width) ;
}
}
this.SetFont("",this.customstyle[xi]);
this.MultiCell(xw,this.cellheight,xdata[xi],this.customborders[xi],this.cellaligns[xi]);
this.SetXY(xx+xw,xy1);
}
this.Ln(xh);
}
this.CheckPageBreak=function(xh)
{
if(this.GetY()+xh>this.PageBreakTrigger)this.AddPage(this.CurOrientation);
}
this.NbLines=function(xw , xtxt)
{
var xnb;
xcw=this.CurrentFont["cw"];
if(xw==0)xw=this.w-(this.rMargin)-this.x;
xwmax=((xw)-2*(this.cMargin))*1000/(this.FontSize);
xs=lib.str_replace("\r","",xtxt);
xnb=xs.length;
if(xnb>0 && xs.charAt(xnb-1)=="\n")xnb--;
xsep=-1;
xi=0;
xj=0;
xl=0;
xnl=1;
while(xi<xnb)
{
xc=xs.charAt(xi);
if(xc=="\n")
{
xi++;
xsep=-1;
xj=xi;
xl=0;
xnl++;
continue;
}
if(xc==" ")xsep=xi;
xl+=(xcw[xc]);
if(xl>xwmax)
{
if(xsep==-1)
{
if(xi==xj)xi++;
}
else xi=xsep+1;
xsep=-1;
xj=xi;
xl=0;
xnl++;
}
else {xi++;}
}
return xnl;
}