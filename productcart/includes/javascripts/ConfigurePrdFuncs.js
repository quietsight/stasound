function open_win(url)
{
    var newWin=window.open(url,'','height=400,width=400,scrollbars=yes,resizable=yes');
    newWin.focus();
}

function win(fileName)
{
	myFloater = window.open('','myWindow','scrollbars=auto,status=no,width=400,height=300')
	myFloater.location.href = fileName;
}


function optwin(fileName)
{
	myFloater = window.open('','myWindow','scrollbars=yes,status=no,width=400,height=300')
	myFloater.location.href = fileName;
}	

function besubmit()
{
	document.getElementById("show_1").style.display = '';;
	return (false);
}

function isDigit(s)
{
	var test=""+s;
	if(test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
	{
		return(true) ;
	}
	return(false);
}

function allDigit(s)
{
	var test=""+s ;
	for (var k=0; k <test.length; k++)
	{
		var c=test.substring(k,k+1);
		if (isDigit(c)==false)
		{
			return (false);
		}
	}
	return (true);
}

function testquantity(fname)
{
	if (eval(fname).value == "0")
	{
		eval(fname).value="1";
	}
}

function testdropdown(fname)
{
	var calValue= eval(fname).value;
	var cal_array=calValue.split("_");
	var mytest=cal_array[0];
	if (mytest == "0")
	{
		eval(fname + "QF").value="0";
	}
	else
	{
		var qtyValue= eval(fname+mytest+"HF").value;
		var qty_array=qtyValue.split("_");
		var minQty=qty_array[1];
		eval(fname + "QF").value=minQty;
	}
}

function testdropqty(qtyname,fname)
{
	var calValue= eval(fname).value;
	var cal_array=calValue.split("_");
	var mytest=cal_array[0];
	if (mytest == "0")
	{
		eval(qtyname).value="0";
		calculate(eval(fname),0);
	}
	else
	{
		if (eval(qtyname).value=="")
		{
			eval(qtyname).value="1"
		}
		if (eval(qtyname).value=="0")
		{
			eval(qtyname).value="1"
		}
		qtyValue= eval(fname+mytest+"HF").value;
		qty_array=qtyValue.split("_");
		QtyValid=qty_array[0];
		minQty=qty_array[1];
		if (qttverify(qtyname,QtyValid,minQty))
		{
			calculate(eval(fname),0);
		}
	}
}