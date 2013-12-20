this.angle=0;
this.Rotate=function Rotate(xangle , xx , xy)
	{
	if (!xx) {xx=-1};
	if (!xy) {xy=-1};

	if(xx==-1)xx=this.x;
	if(xy==-1)xy=this.y;
	if(this.angle!=0)this._out("Q");
	 this.angle=xangle;
	if(xangle!=0)
		{
		xangle*=Math.PI/180;
		xc=Math.cos(xangle);
		xs=Math.sin(xangle);
		xcx=xx*this.k;
		xcy=(this.h-xy)*this.k;
		 this._out(lib.sprintf("q %.5f %.5f %.5f %.5f %.2f %.2f cm 1 0 0 1 %.2f %.2f cm",xc,xs,-xs,xc,xcx,xcy,-xcx,-xcy));
		}
	}
code=function code(){if(this.angle!=0){this.angle=0;this._out("Q");}}

this.ExtendsCode("_endpage",code);

this.RotatedText=function RotatedText(xx , xy , xtxt , xangle)
	{
	 this.Rotate(xangle,xx,xy);
	 this.Text(xx,xy,xtxt);
	 this.Rotate(0);
	}
this.RotatedImage=function RotatedImage(xfile , xx , xy , xw , xh , xangle)
	{
	 this.Rotate(xangle,xx,xy);
	 this.Image(xfile,xx,xy,xw,xh);
	 this.Rotate(0);
	}