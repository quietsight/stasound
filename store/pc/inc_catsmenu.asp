<%if session("customerType")="1" then%>
	<!--#include file="inc_WholeSaleCatMenu.inc"-->
<%else%>
	<!--#include file="inc_RetailCatMenu.inc"-->
<%end if%>
<script>
var imgopen = new Image();
imgopen.src = "images/btn_collapse.gif";
var imageclose = new Image();
imageclose.src = "images/btn_expand.gif";

function UpDown(tabid)
{
	try
	{
		var etab=document.getElementById('SUB' + tabid);
		if (etab.style.display=='')
		{
			etab.style.display='none';
			var etab=document.images['IMGCAT' + tabid];
			etab.src=imageclose.src;
		}
		else
		{	
			etab.style.display='';
			var etab=document.images['IMGCAT' + tabid];
			etab.src=imgopen.src;
		}
	}
	catch(err)
	{
		return(false);
	}
	
	
}

<%
Function ExpandParent(pcv_IDCAT)
Dim tmp_query,rsTest
	query="SELECT idparentCategory FROM categories where idcategory=" & pcv_IDCAT
	set rsTest=conlayout.execute(query)
	if not rsTest.eof then
		pcv_tmpParent=rsTest("idparentCategory")
		set rsTest=nothing
		if pcv_tmpParent<>"0" and pcv_tmpParent<>"1" then%>
			UpDown(<%=pcv_tmpParent%>);
			<%call ExpandParent(pcv_tmpParent)
		end if
	end if
	set rsTest=nothing
End Function

	if validNum(pIdCategory) and (pIdCategory<>"0") then
	call ExpandParent(pIdCategory)%>
	UpDown(<%=pIdCategory%>);
	<%
	end if
	%>
</script>
