var FormArr=new Array();
var FieldCount=0;

var FormItem1=new Array();
var FormItem2=new Array();
var FormItemCount=-1;

var FormDrop1=new Array();
var FormDrop2=new Array();
var FormDrop3=new Array();
var FormDropCount=-1;

var FormRadio1=new Array();
var FormRadio2=new Array();
var FormRadio3=new Array();
var FormRadio4=new Array();
var FormRadioCount=-1;

var FormCB1=new Array();
var FormCB2=new Array();
var FormCB3=new Array();
var FormCBCount=-1;

var pcv_ResetForm=0;

//Find Item Location in the Form
function GetItemLocation()
{
var i=0;
var k=0;
var m=0;
var tmpRadioName="";
var tmpCBName="";
var objElems = document.additem.elements;
var j=objElems.length;
do
{
	i=j-1;
	var tmptype=objElems[i].type;
	if (tmptype=="select-one")
	{
			FormDropCount=FormDropCount+1;
			FormDrop1[FormDropCount]=i;
			FormDrop2[FormDropCount]=objElems[i].name;
			FormDrop3[FormDropCount]=-1;
			var m=objElems[i].options.length;
			do
			{
				k=m-1;
				var tmpStr1=objElems[i].options[k].value;
				var tmpStr2=tmpStr1.split("_");
				var tmpid=tmpStr2[0];
				if (eval(tmpStr2[0])!=0)
				{
					FormItemCount=FormItemCount+1;
					FormItem1[FormItemCount]=tmpStr2[0];
					FormItem2[FormItemCount]=i + "_1_" + k + "_" + FormDropCount;
				}
			}
			while (--m);
	}
	else
	{
		if (tmptype=="radio")
		{
			var tmpStr1=objElems[i].value;
			var tmpStr2=tmpStr1.split("_");
			if (objElems[i].name!=tmpRadioName)
			{
				FormRadioCount=FormRadioCount+1;
				tmpRadioName=objElems[i].name;
				FormRadio1[FormRadioCount]=tmpRadioName;
				FormRadio2[FormRadioCount]=i;
				FormRadio3[FormRadioCount]=i;
				FormRadio4[FormRadioCount]=0;
			}
			else
			{
				//if (eval(tmpStr2[0])!=0)
				FormRadio2[FormRadioCount]=i;
			}
			if (eval(tmpStr2[0])!=0)
			{
				FormItemCount=FormItemCount+1;
				FormItem1[FormItemCount]=tmpStr2[0];
				FormItem2[FormItemCount]="0_0_" + i + "_" + FormRadioCount;
			}
		}
		else
		{
		if (tmptype=="checkbox")
		{
			var tmpStr1=objElems[i].value;
			var tmpStr2=tmpStr1.split("_");
			var tmpname=objElems[i].name;
			if ((tmpCBName=="") || (tmpname.indexOf(tmpCBName)!=0))
			{
				FormCBCount=FormCBCount+1;
				var tmpname1=tmpStr2[0] + "";
				tmpCBName=tmpname.substr(0,tmpname.length-tmpname1.length)
				FormCB1[FormCBCount]=tmpCBName;
				FormCB2[FormCBCount]=i;
				FormCB3[FormCBCount]=i;
			}
			else
			{
				if ((tmpCBName!="") && (tmpname.indexOf(tmpCBName)==0)) FormCB2[FormCBCount]=i;
			}
			if (eval(tmpStr2[0])!=0)
			{
				FormItemCount=FormItemCount+1;
				FormItem1[FormItemCount]=tmpStr2[0];
				FormItem2[FormItemCount]="0_2_" + i + "_" + FormCBCount;
			}
		}
		}
	}
}
while (--j);
}

//Find Item Location in the Saved Array
function New_FindItemLocation(itemID)
{
	var i=FormItemCount+1;
	var j=Math.round((FormItemCount+1)/2);
	var m=-1;
	do
	{
		i--;
		if (eval(FormItem1[i])==eval(itemID))
		{
			return(FormItem2[i]);
			break;
		}
		m++;
		if (FormItem1[m]==itemID)
		{
			return(FormItem2[m]);
			break;
		}
	}
	while (--j);
}

function New_FindItemInRadioList(tmpindex,tvalue)
{
	var m=FormRadio2[tmpindex];
	var k=FormRadio3[tmpindex];
	var tmp1=0;
	var i=-1;
	for (tmp1=m; tmp1<=k;tmp1++)
	{
		if (document.additem.elements[tmp1].type=="radio")
		{
			i=i+1;
			var tmpvalue=document.additem.elements[tmp1].value;
			if (tmpvalue.indexOf(tvalue + "_")==0)
			{
				return(i);
				break;
			}
		}
	}
	return(-1);
}

function New_GetField()
{
	var i=0;
	var tmpRadioName="";
	var objElems = document.additem.elements;
	i=-1;
	var j=objElems.length-1;
	do
	{
		i++;
		var tmptype=objElems[i].type;
		if (tmptype=="select-one")
		{
			FieldCount=FieldCount+1;
			FormArr[FieldCount-1]=objElems[i].name;
		}
		else
		{
			if (tmptype=="radio")
			{
				if (objElems[i].name!=tmpRadioName)
				{
					FieldCount=FieldCount+1;
					tmpRadioName=objElems[i].name;
					FormArr[FieldCount-1]=tmpRadioName;
				}
			}
			else
			{
				if (tmptype=="checkbox")
				{
					FieldCount=FieldCount+1;
					FormArr[FieldCount-1]=objElems[i].name;
				}
			}
		}
	}
	while (i<j);
}

function New_GetFieldIndex(fieldname)
{
	var i=0;
	var j=FieldCount;
	do
	{
		i=j-1;
		if (FormArr[i]==fieldname)
		{
			return(i+1);
			break;
		}
	}
	while (--j);
}

function GenDropInfo(xfield)
{
	GenDropListInfo(xfield,xfield.selectedIndex,0,1,xfield.selectedIndex);
}

function GenDropListInfo(xfield,tindex,defprice,firstitem,defindex)
{
	var saveprice=0;
	if (pcv_ResetForm==1)
	{
		xfield.options[tindex].style.color="";
	}
	var tempStr1=xfield.options[tindex].text;
	var str_array=tempStr1.split(" - " + optmsg1);
	var str1_array=str_array[0].split(" - " + optmsg2);
	var tempStr=str1_array[0];
	var calPrice1=xfield.options[tindex].value;										
	var cal_array=calPrice1.split("_");
	var IDProduct=cal_array[0];
	var CalSPrice=parseFloatEx(cal_array[1])-parseFloatEx(defprice);
	
	var CalPrice=cal_array[3];
	var CalQty=eval("document.additem." + xfield.name + "QF").value;
	if (firstitem!=1) CalQty=new_getDefMinQty(cal_array[0],0);
	if (CalQty=="0") { CalQty="1" };
	var testAC=0;
	try
	{
		document.additem.CHGTotal.value =0;
	}
	catch(err)
	{
		testAC=1;
	}
	if (testAC==1)
	{
		var PrdQty=eval("document.additem.quantity.value");
	}
	else
	{
		var PrdQty=1;
	}
	if (firstitem==1)
	{
		if ((cal_array[0]!=0) && (cal_array[1]==0))
		{
			CalSPrice=eval(CalSPrice)+(eval(CalQty)-new_getDefMinQty(cal_array[0],1))*eval(CalPrice);
		}
		else
		{
			CalSPrice=eval(CalSPrice)+(eval(CalQty)-1)*eval(CalPrice)-parseFloatEx(DisValue(IDProduct,(CalQty)*PrdQty,CalPrice)/PrdQty);
		}
		saveprice=CalSPrice;
	}
	if (firstitem==1)
	{
		var NewPrice=0;
	}
	else
	{
		if ((cal_array[0]!=0) && (cal_array[1]==0))
		{
			var NewPrice=eval(CalSPrice*PrdQty)+(eval(CalQty)-new_getDefMinQty(cal_array[0],1))*eval(PrdQty)*eval(CalPrice)-DisValue(IDProduct,(CalQty)*PrdQty,CalPrice);
		}
		else
		{
			var NewPrice=eval(CalSPrice*PrdQty)+(eval(CalQty)-1)*eval(PrdQty)*eval(CalPrice)-DisValue(IDProduct,(CalQty)*PrdQty,CalPrice);
		}
	}
	if (showprices<2)
	{
		if (NewPrice > 0 )
		{
			var tempStr=tempStr + " - " + optmsg1;
		}
		else
		{
			if (NewPrice < 0)
			{
				var tempStr=tempStr + " - " + optmsg2;
				var NewPrice=eval(-1*NewPrice);
			}
		}
		if (NewPrice > 0 )
		{
			var tempStr2=New_FormatNumber(NewPrice);
		}
		else
		{
			tempStr2="";
		}
		xfield.options[tindex].text=tempStr + tempStr2;
	}
	else
	{
		xfield.options[tindex].text=tempStr;
	}
	
	if (firstitem==1)
	{
		var totaldrop=eval("document.additem." + xfield.name +".options").length;
		if (parseInt(totaldrop)>0)
		{
			var i=0;
			for (i=0;i<totaldrop;i++)
			{
				if (i!=defindex) GenDropListInfo(xfield,i,saveprice,0,0);
			}
		}
	}
}

function GenRadioInfo(xfield,tindex)
{
	GenRadioListInfo(xfield,tindex,0,1,tindex);
}

function GenRadioExtInfo(xfield)
{
	tindex=-1;
	var totalradio=xfield.length;
	if (parseInt(totalradio)>0)
	{
		var i=0;
		for (i=0;i<totalradio;i++)
		{
			if (xfield[i].checked==true)
			{
				tindex=i;
				break;
			}
		}
	}
	if (tindex>=0) { GenRadioListInfo(xfield[0],tindex,0,1,tindex)}
	else
	{
		if (pcv_ResetForm==1)
		{
			if (parseInt(totalradio)>0)
			{
				tindex=-2;
				GenRadioListInfo(xfield[0],tindex,0,1,tindex);
			}
			else
			{
				if (xfield.checked==true)
				{
					tindex=-1;
					GenRadioListInfo(xfield,tindex,0,1,tindex);
				}
				else
				{
					tindex=-1;
					GenRadioListInfo(xfield,tindex,0,0,tindex);
				}
			}
		}
	}
}

function GenRadioListInfo(xfield,tindex,defprice,firstitem,defindex)
{
	var saveprice=0;
	
	if (tindex!=-2)
	{
		if (tindex>=0)
		{
			var calPrice1=eval("document.additem." + xfield.name + "[" + tindex + "]").value;
			if (pcv_ResetForm==1)
			{
				eval("document.additem." + xfield.name + "[" + tindex + "]").disabled=false;
			}
		}
		else
		{
			var calPrice1=xfield.value;
			if (pcv_ResetForm==1)
			{
				xfield.disabled=false;
			}
		}
		var cal_array=calPrice1.split("_");
		var IDProduct=cal_array[0];
		var CalSPrice=parseFloatEx(cal_array[1])-parseFloatEx(defprice);
		var CalPrice=cal_array[3];
		if (tindex>=0)
		{
			var CalQty=eval("document.additem." + xfield.name + "QF" + tindex).value;
		}
		else
		{
			var CalQty=eval("document.additem." + xfield.name + "QF0").value;
		}
		if (CalQty=="0")
		{
			CalQty=new_getDefMinQty(cal_array[0],0);
		}
		if (CalQty=="0") CalQty="1";
		var testAC=0;
		try
		{
			document.additem.CHGTotal.value =0;
		}
		catch(err)
		{
			testAC=1;
		}
		if (testAC==1)
		{
			var PrdQty=eval("document.additem.quantity.value");
		}
		else
		{
			var PrdQty=1;
		}
		if (firstitem==1)
		{
			if ((cal_array[0]!=0) && (cal_array[1]==0))
			{
				CalSPrice=eval(CalSPrice)+(eval(CalQty)-new_getDefMinQty(cal_array[0],1))*eval(CalPrice);
			}
			else
			{
				CalSPrice=eval(CalSPrice)+(eval(CalQty)-1)*eval(CalPrice)-parseFloatEx(DisValue(IDProduct,(CalQty)*PrdQty,CalPrice)/PrdQty);
			}
			saveprice=CalSPrice;
		}
		if (firstitem==1)
		{
			var NewPrice=0;
		}
		else
		{
			if ((cal_array[0]!=0) && (cal_array[1]==0))
			{
				var NewPrice=eval(CalSPrice*PrdQty)+(eval(CalQty)-new_getDefMinQty(cal_array[0],1))*eval(PrdQty)*eval(CalPrice)-DisValue(IDProduct,(CalQty)*PrdQty,CalPrice);
			}
			else
			{
				var NewPrice=eval(CalSPrice*PrdQty)+(eval(CalQty)-1)*eval(PrdQty)*eval(CalPrice)-DisValue(IDProduct,(CalQty)*PrdQty,CalPrice);
			}
		}
		if (NewPrice > 0 )
		{
			var tempStr=" - " + optmsg1;
		}
		else
		{
			if (NewPrice < 0)
			{
				var tempStr=" - " + optmsg2;
				var NewPrice=eval(-1*NewPrice);
			}
			else
			{
				var tempStr=" ";
			}
		}
		if (NewPrice > 0 )
		{
			tempStr2 =New_FormatNumber(NewPrice);
		}
		else
		{
			tempStr2="";
		}
		if (tindex>=0)
		{
			var myStr=eval("document.additem." + xfield.name + "TX" + tindex);
		}
		else
		{
			var myStr=eval("document.additem." + xfield.name + "TX0");
		}
		if (showprices<2)
		{ 
			myStr.value=tempStr + tempStr2;
		}
		else
		{
			myStr.value=" ";
		}
		var mStr=tempStr+tempStr2;
		myStr.size=mStr.length;
	} //Have default item
	
	if ((firstitem==1) && (tindex!=-1))
	{
		var totalradio=eval("document.additem." + xfield.name).length;
		if (parseInt(totalradio)>0)
		{
			var i=0;
			for (i=0;i<totalradio;i++)
			{
				if (i!=defindex) GenRadioListInfo(xfield,i,saveprice,0,0);
			}
		}
	}
}

function New_AutoUpdateQtyPrice()
{
	if (eval("document.additem.quantity.value")!=eval("document.additem.savequantity.value"))
	{
	document.additem.savequantity.value=document.additem.quantity.value;
	var i=0;
	var objElems = document.additem.elements;
	var j=objElems.length;
	do
	{
		i=j-1;
		var tmptype=document.additem.elements[i].type;
		if (tmptype=="select-one")
		{
			calculate(document.additem.elements[i],1)
		}
		else
		{
			if (tmptype=="radio")
			{
				if (document.additem.elements[i].checked==true)	calculate(document.additem.elements[i],1);
			}
			else
			{
				if (tmptype=="checkbox")
				{
					if (document.additem.elements[i].checked==true)
					{
						calculate(document.additem.elements[i],1);
					}
				}
			}
		}
	}
	while (--j);
	New_calculateAll();
	}
}

function new_getDefMinQty(tmpID,tmpdef)
{
var tmp=1;
var i=0;

	
	for (i=0;i<=defitemscount;i++)
	{
		if (tmpdef==1)
		{
			if ((defitems[i] + ""==tmpID + "") && (defset[i]==tmpdef))
			{
				tmp=defmin[i];
				if (tmp==0) tmp=1;
				break;
			}
		}
		else
		{
			if (defitems[i] + ""==tmpID + "")
			{
				tmp=defmin[i];
				if (tmp==0) tmp=1;
				break;
			}
		}
	}
	return(tmp)
}


function calculate(xfield,ctype)
{
	var testAC=0;
	try
	{
		document.additem.CHGTotal.value =0;
	}
	catch(err)
	{
		testAC=1;
	}
	if (testAC==1)
	{
		var PQty=eval("document.additem.quantity.value");
	}
	else
	{
		var PQty=1;
	}

	if (ctype==2)
	{
		if (xfield.length>0)
		{
			var xfield1=xfield[0];
		}
		else
		{
			var xfield1=xfield;
		}
		var tmptype=xfield1.type;
	}
	else
	{
		var tmptype=xfield.type;
		var xfield1=xfield;
	}
	if (tmptype=="select-one")
	{
			//Drop-down
			if (pcv_ResetForm==0) GenDropInfo(xfield);
			var tmpname=xfield.name;
			var calPrice=xfield.value;
			var cal_array=calPrice.split("_");
			var tmpobj=eval("document.additem." + tmpname + "QF");
			if (eval(cal_array[0]) == 0)
			{
				tmpobj.value="1";
				var tmpindex=New_GetFieldIndex(tmpname);
				QD1=QD1-parseFloatEx(eval("document.additem.Discount" + tmpindex).value);
				QD1=Math.round(QD1*100)/100;
				Ctotal=Ctotal-parseFloatEx(eval("document.additem.currentValue" + tmpindex).value);
				Ctotal=Math.round(Ctotal*100)/100;
				eval("document.additem.Discount" + tmpindex).value=0;
				eval("document.additem.currentValue" + tmpindex).value =eval(cal_array[1]);
				Ctotal=Ctotal+parseFloatEx(eval("document.additem.currentValue" + tmpindex).value);
			}
			else
			{
				if (tmpobj.value==0) tmpobj.value=1;
				CalQty=tmpobj.value-new_getDefMinQty(cal_array[0],1);
				if (CalQty >=0)
				{
					CalQPrice=CalQty*eval(cal_array[3]);
				}
				else
				{
					CalQPrice=0;
				}
				var tmpindex=New_GetFieldIndex(tmpname);
				QD1=QD1-parseFloatEx(eval("document.additem.Discount" + tmpindex).value);
				QD1=Math.round(QD1*100)/100;
				Ctotal=Ctotal-parseFloatEx(eval("document.additem.currentValue" + tmpindex).value);
				Ctotal=Math.round(Ctotal*100)/100;
				eval("document.additem.Discount" + tmpindex).value=DisValue(cal_array[0],(CalQty+new_getDefMinQty(cal_array[0],1))*PQty,cal_array[3]);
				eval("document.additem.currentValue" + tmpindex).value = eval(cal_array[1])+CalQPrice;
				QD1=QD1+parseFloatEx(eval("document.additem.Discount" + tmpindex).value);
				Ctotal=Ctotal+parseFloatEx(eval("document.additem.currentValue" + tmpindex).value);
			}
			if (ctype==1) return(true);
	}
	else
	{
		if (tmptype=="radio")
		{
			//Radio
			var tmpname=xfield1.name;
			var iEleCnt=-1;
			if (eval("document.additem." + tmpname).length>0)
			{
				var i=0;
				var j=eval("document.additem." + tmpname).length;
				for (i=j-1;i>=0;i--)
				{
					if (eval("document.additem." + tmpname +"[" + i + "]").checked==true)
					{
						var iEleCnt=i;
					}
					else
					{
						eval("document.additem." + tmpname + "QF" + i).value=0;
					}
				}
				if (iEleCnt!=-1)
				{
					if (pcv_ResetForm==0) GenRadioInfo(eval("document.additem." + tmpname +"[" + iEleCnt + "]"),iEleCnt);
					xfield1=eval("document.additem." + tmpname +"[" + iEleCnt + "]");
					var tmpStr1=xfield1.value;
					var tmpStr2=tmpStr1.split("_");
				
					if (tmpStr2[0]!=0)
					{
						try
						{
						testquantity("document.additem." + tmpname + "QF" + iEleCnt);
						}
						catch(err){}
						var calPrice= xfield1.value;
						var cal_array=calPrice.split("_");
						eval("document.additem." + tmpname + "QF").value=eval("document.additem." + tmpname + "QF" + iEleCnt).value;
						CalQty=eval("document.additem." + tmpname + "QF" +iEleCnt).value-new_getDefMinQty(cal_array[0],1);
						if (CalQty >=0)
						{
							CalQPrice=CalQty*eval(cal_array[3]);
						}
						else
						{
							CalQPrice=0
						}
						var tmpindex=New_GetFieldIndex(tmpname);
						QD1=QD1-parseFloatEx(eval("document.additem.Discount" + tmpindex).value);
						QD1=Math.round(QD1*100)/100;
						Ctotal=Ctotal-parseFloatEx(eval("document.additem.currentValue" + tmpindex).value);
						Ctotal=Math.round(Ctotal*100)/100;
						eval("document.additem.Discount" + tmpindex).value=DisValue(cal_array[0],(CalQty+new_getDefMinQty(cal_array[0],1))*PQty,cal_array[3]);
						eval("document.additem.currentValue" + tmpindex).value = eval(cal_array[1])+CalQPrice;
						QD1=QD1+parseFloatEx(eval("document.additem.Discount" + tmpindex).value);
						Ctotal=Ctotal+parseFloatEx(eval("document.additem.currentValue" + tmpindex).value);
					}
					else
					{
						eval("document.additem." + tmpname + "QF" + iEleCnt).value="1";
						var tmpindex=New_GetFieldIndex(tmpname);
						QD1=QD1-parseFloatEx(eval("document.additem.Discount" + tmpindex).value);
						QD1=Math.round(QD1*100)/100;
						Ctotal=Ctotal-parseFloatEx(eval("document.additem.currentValue" + tmpindex).value);
						Ctotal=Math.round(Ctotal*100)/100;
						eval("document.additem.Discount" + tmpindex).value=0;
						eval("document.additem.currentValue" + tmpindex).value =eval(tmpStr2[1]);
						Ctotal=Ctotal+parseFloatEx(eval("document.additem.currentValue" + tmpindex).value);
					}
				} //Have selected item
			}
			else
			{
				if (xfield1.checked==true)
				{
					if (pcv_ResetForm==0) GenRadioInfo(xfield1,-1);
					try
					{
					testquantity("document.additem." + xfield1.name + "QF0");
					}
					catch(err){}
					var calPrice= xfield1.value;
					var cal_array=calPrice.split("_");
					eval("document.additem." + tmpname + "QF").value=eval("document.additem." + xfield1.name + "QF0").value;
					CalQty=eval("document.additem." + xfield1.name + "QF0").value-new_getDefMinQty(cal_array[0],1);
					if (CalQty >=0)
					{
						CalQPrice=CalQty*eval(cal_array[3]);
					}
					else
					{
						CalQPrice=0
					}
					var tmpindex=New_GetFieldIndex(tmpname);
					QD1=QD1-parseFloatEx(eval("document.additem.Discount" + tmpindex).value);
					QD1=Math.round(QD1*100)/100;
					Ctotal=Ctotal-parseFloatEx(eval("document.additem.currentValue" + tmpindex).value);
					Ctotal=Math.round(Ctotal*100)/100;
					eval("document.additem.Discount" + tmpindex).value=DisValue(cal_array[0],(CalQty+new_getDefMinQty(cal_array[0],1))*PQty,cal_array[3]);
					eval("document.additem.currentValue" + tmpindex).value = eval(cal_array[1])+CalQPrice;
					QD1=QD1+parseFloatEx(eval("document.additem.Discount" + tmpindex).value);
					Ctotal=Ctotal+parseFloatEx(eval("document.additem.currentValue" + tmpindex).value);
				}
				//Have selected item
			}
			if (ctype==1) return(true);
		}
		else
		{
		if (tmptype=="checkbox")
		{
			//Checkbox
			if (xfield.checked == true)
			{
				var tmpname=xfield.name;
				try
				{
				testquantity("document.additem." + tmpname + "QF");
				}
				catch(err){}
				var calPrice=xfield.value;
				var cal_array=calPrice.split("_");
				CalQty=eval("document.additem." + tmpname +"QF").value-new_getDefMinQty(cal_array[0],1);
				if (CalQty >=0)
				{
					CalQPrice=CalQty*eval(cal_array[3]);
				}
				else
				{
					CalQPrice=0;
				}
				var tmpindex=New_GetFieldIndex(tmpname);
				QD1=QD1-parseFloatEx(eval("document.additem.Discount" + tmpindex).value);
				QD1=Math.round(QD1*100)/100;
				Ctotal=Ctotal-parseFloatEx(eval("document.additem.currentValue" + tmpindex).value);
				Ctotal=Math.round(Ctotal*100)/100;
				eval("document.additem.Discount" + tmpindex).value=DisValue(cal_array[0],(CalQty+new_getDefMinQty(cal_array[0],1))*PQty,cal_array[3]);
				eval("document.additem.currentValue" + tmpindex).value = eval(cal_array[1])+CalQPrice;
				QD1=QD1+parseFloatEx(eval("document.additem.Discount" + tmpindex).value);
				Ctotal=Ctotal+parseFloatEx(eval("document.additem.currentValue" + tmpindex).value);
			}
			else
			{
				var tmpname=xfield.name;
				eval("document.additem."+ tmpname + "QF").value=0;
				var calPrice=xfield.value;
				var cal_array=calPrice.split("_");
				var tmpindex=New_GetFieldIndex(tmpname);
				QD1=QD1-parseFloatEx(eval("document.additem.Discount" + tmpindex).value);
				QD1=Math.round(QD1*100)/100;
				Ctotal=Ctotal-parseFloatEx(eval("document.additem.currentValue" + tmpindex).value);
				Ctotal=Math.round(Ctotal*100)/100;
				eval("document.additem.Discount" + tmpindex).value=0;
				if ((parseFloatEx(cal_array[3])>0) && (parseFloatEx(cal_array[1])==0))
				{
					Ctotal=Ctotal-parseFloatEx(Math.round(cal_array[3]*new_getDefMinQty(cal_array[0],1)*100)/100);
					Ctotal=Math.round(Ctotal*100)/100;
					eval("document.additem.currentValue" + tmpindex).value='-' + Math.round(cal_array[3]*new_getDefMinQty(cal_array[0],1)*100)/100;
				}
				else
				{
					eval("document.additem.currentValue" + tmpindex).value=0;
				}
			}
			if (ctype==1) return(true);
		}
		}

	}
	New_calculateAll();
}

function New_FormatNumber(tmpvalue)
{
	var DifferenceTotal = new DNumberFormat();
	DifferenceTotal.setNumber(tmpvalue);
	if (scDecSign==",")
	{
		DifferenceTotal.setSeparators(true,DifferenceTotal.PERIOD);
	}
	else
	{
		DifferenceTotal.setCommas(true);
	}
	DifferenceTotal.setPlaces(2);
	DifferenceTotal.setCurrency(true);
	DifferenceTotal.setCurrencyPrefix(scCurSign);
	return(DifferenceTotal.toFormatted());
}

function New_calculateAll()
{
	var testAC=0;
	try
	{
		document.additem.CHGTotal.value =0;
	}
	catch(err)
	{
		testAC=1;
	}
	if (testAC==1)
	{
		var PQty=eval("document.additem.quantity.value");
	}
	else
	{
		var PQty=1;
	}
	var DifferenceTotal=New_FormatNumber(Ctotal);
	document.additem.total.value = DifferenceTotal;
	TL1=Ctotal*PQty;
	TL2=parseFloatEx(eval("document.additem.currentValue0.value"))*PQty;
	document.additem.TLPriceDefault.value=TL2;
	var TLDTotal=New_FormatNumber(parseFloatEx(eval("document.additem.currentValue0.value"))*PQty);
	document.additem.TLcurPrice.value = TLDTotal;
	document.additem.TLdefaultprice.value = TLDTotal;

	var TLCTotal=New_FormatNumber(Ctotal*PQty);
	document.additem.TLtotal.value = TLCTotal;
	
	QD2=-1*QD1;
	var QD=New_FormatNumber(QD2);
	TL3=QD2;
	document.additem.Discounts.value = QD;
				
	TL5=0;

	TL5=QDisValue(tmpIDProduct,document.additem.quantity,TL1+TL2+TL3); //Using Configured Price to calculate Product Quantity Discount
	
	try {
		if (pcQDiscountType=="1")
		{
			TL5=QDisValue(tmpIDProduct,document.additem.quantity,TL2); //Using Default Price to calculate Product Quantity Discount
		}
	}
	catch(err){}

	document.additem.QDiscounts0.value =TL5;
	var QtyD=New_FormatNumber(-1*TL5);
	document.additem.QDiscounts.value = QtyD;
	
	var tmpResult=TL1+TL2+TL3-TL5;
	tmpResult=Math.round(tmpResult*100)/100;
	var TLQtyD=New_FormatNumber(tmpResult);
	document.additem.TotalWithQD.value = TLQtyD;
	document.additem.TLGrandTotal2QD.value = TLQtyD;
	
	var GrandTotal1=New_FormatNumber(Ctotal+parseFloatEx(eval("document.additem.currentValue0.value")));
	document.additem.GrandTotal.value = GrandTotal1;
				
	document.additem.CMDefault.value=TL1+TL3;
	var tmpResult=TL1+TL3-TL5;
	tmpResult=Math.round(tmpResult*100)/100;
	document.additem.CMWQD.value=tmpResult;
				
	var TLtotal=New_FormatNumber(TL1+TL2+TL3);
	document.additem.TLGrandTotal.value = TLtotal;
	document.additem.TLGrandTotal2.value = TLtotal;

	document.additem.GrandTotal2.value = GrandTotal1;
	
	var testAC=0;
	try
	{
		document.additem.CHGTotal.value =0;
	}
	catch(err)
	{
		testAC=1;
	}
	if (testAC==0)
	{
		document.additem.total.value = New_FormatNumber(parseFloatEx(Ctotal)+parseFloatEx(TL3));
		document.additem.CHGTotal.value = parseFloatEx(Ctotal)+parseFloatEx(TL3);
		var GrandTotal1=New_FormatNumber(parseFloatEx(eval("document.additem.CMWQD0.value"))+Ctotal+parseFloatEx(TL3)+parseFloatEx(eval("document.additem.currentValue0.value")));
		document.additem.GrandTotal.value = GrandTotal1;
		document.additem.GrandTotalQD.value = GrandTotal1;
	}
	else
	{
		CheckTotalItemQty();
	}
}


//Reset From Field properties
function new_ResetProperties()
{
pcv_ResetForm=1;
var i=0;
var objElems = document.additem.elements;
var j=objElems.length;
var tmpname="";
do
{
	i=j-1;
	var tmptype=objElems[i].type;
	if (tmptype=="select-one")
	{
		GenDropInfo(objElems[i]);
		calculate(objElems[i],1);
	}
	else
	{
		if (tmptype=="radio")
		{
			if (objElems[i].name!=tmpname)
			{
				var tmpname=objElems[i].name;
				GenRadioExtInfo(eval("document.additem." + tmpname));
				calculate(eval("document.additem." + tmpname),3);
			}
		}
		else
		{
			if (tmptype=="checkbox")
			{
				objElems[i].disabled=false;
				calculate(objElems[i],1);
			}
		}
	}
}
while (--j);
pcv_ResetForm=0;
}

//Generate Infor OnLoad
function new_GenInforOnLoad()
{
var i=0;
var objElems = document.additem.elements;
var j=objElems.length;
var tmpname="";
do
{
	i=j-1;
	var tmptype=objElems[i].type;
	if (tmptype=="select-one")
	{
		GenDropInfo(objElems[i]);
	}
	else
	{
		if (tmptype=="radio")
		{
			if (objElems[i].name!=tmpname)
			{
				var tmpname=objElems[i].name;
				GenRadioExtInfo(eval("document.additem." + tmpname));
			}
		}
	}
}
while (--j);
}

function parseFloatEx(tmpvalue)
{
	var tmp1=""+tmpvalue;
	if (scDecSign==",")	tmp1=tmp1.replace(",",".");
	return(parseFloat(tmp1));
}

function CheckTotalItemQty()
{
var i=0;
var k=0;
var m=0;
var TotalItemQty=0;
var tmpRadioName="";
var tmpCBName="";
if ((TotalMaxSelect>0))
{
var objElems = document.additem.elements;
var j=objElems.length;
for (i=j-1;i>=0;i--)
{
	var tmptype=objElems[i].type;
	if (tmptype=="select-one")
	{
			var tmpname=objElems[i].name;
			TotalItemQty=TotalItemQty+Number(eval("document.additem." + tmpname + "QF").value);
	}
	else
	{
		if (tmptype=="radio")
		{
			if (objElems[i].name!=tmpRadioName)
			{
				var tmpname=objElems[i].name;
				var tmpRadioName=objElems[i].name;
				var iEleCnt=-1;
				if (eval("document.additem." + tmpname).length>0)
				{
					var k=0;
					var m=eval("document.additem." + tmpname).length;
					for (k=m-1;k>=0;k--)
					{
						if (eval("document.additem." + tmpname +"[" + k + "]").checked==true)
						{
							var iEleCnt=k;
						}
						else
						{
							eval("document.additem." + tmpname + "QF" + k).value=0;
						}
					}
					if (iEleCnt!=-1)
					{
						xfield1=eval("document.additem." + tmpname +"[" + iEleCnt + "]");
						var tmpStr1=xfield1.value;
						var tmpStr2=tmpStr1.split("_");
				
						if (tmpStr2[0]!=0)
						{
							TotalItemQty=TotalItemQty+Number(eval("document.additem." + tmpname + "QF" + iEleCnt).value);
						}
					}
				}
				else
				{
					if (xfield1.checked==true)
					{
						TotalItemQty=TotalItemQty+Number(eval("document.additem." + xfield1.name + "QF0").value);
					}
				}
			}
		}
		else
		{
			if (tmptype=="checkbox")
			{
				var tmpname=objElems[i].name;
				TotalItemQty=TotalItemQty+Number(eval("document.additem." + tmpname +"QF").value);
			}
		}
	}
}
if (TotalItemQty>TotalMaxSelect)
{
	alert(MaxSelectMsg1 + TotalMaxSelect + MaxSelectMsg1a);
	return(false);
}
} //TotalMaxSelect>0
return(true);
}