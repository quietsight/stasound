<%
pcv_cats1=""
pcv_strTmpCat=""

call opendb()

Function genCatInfor(parentNode,tmp_IDCAT)
Dim k,pcv_first,pcv_prdcount,pcv_back,pcv_end,tsT,tmp_query,pcv_CatName,pcv_prds1,tmp1,tmp2
Dim catNode

	For k=0 to intCountk
		if clng(pcv_cats1(2,k))=clng(tmp_IDCAT) then
			Set catNode=oXML.createNode(1,"Category","")
			catNode.setAttribute "catName",Server.HTMLEncode(ClearHTMLTags2(pcv_cats1(1,k),0))
			catNode.setAttribute "catID",(pcv_cats1(0,k))
			catNode.setAttribute "catParentID",(pcv_cats1(2,k))
			catNode.setAttribute "RetailHide",(pcv_cats1(3,k))
			catNode.setAttribute "iBTOHide",(pcv_cats1(4,k))
			parentNode.appendChild(catNode)
			call genCatInfor(catNode,pcv_cats1(0,k))
		end if
	Next
End Function

Function GenFileName()
	dim fname
	fname="XMLCat-"
	systime=now()
	fname= fname & cstr(year(systime)) & cstr(month(systime)) & cstr(day(systime)) & "-"
	fname= fname  & cstr(hour(systime)) & cstr(minute(systime)) & cstr(second(systime))
	GenFileName=fname
End Function

Set oXML=Server.CreateObject("MSXML2.DOMDocument"&scXML)
Set oRoot = oXML.createNode(1,"Categories","")
oXML.appendChild(oRoot)

query="SELECT idCategory,categoryDesc,idParentCategory,pccats_RetailHide,iBTOhide FROM categories WHERE categories.idCategory<>1 ORDER BY categories.idParentCategory ASC,categories.priority ASC,categories.categoryDesc ASC;"
set rsC=connTemp.execute(query)
if not rsC.eof then
	pcv_cats1=rsC.GetRows()
	intCountk=ubound(pcv_cats1,2)
	set rsC=nothing
	For ik=0 to intCountk
		if pcv_cats1(2,ik)="1" then
			Set subNode=oXML.createNode(1,"Category","")
			subNode.setAttribute "catName",Server.HTMLEncode(ClearHTMLTags2(pcv_cats1(1,ik),0))
			subNode.setAttribute "catID",(pcv_cats1(0,ik))
			subNode.setAttribute "catParentID",(pcv_cats1(2,ik))
			subNode.setAttribute "RetailHide",(pcv_cats1(3,ik))
			subNode.setAttribute "iBTOHide",(pcv_cats1(4,ik))
			oRoot.appendChild(subNode)
			call genCatInfor(subNode,pcv_cats1(0,ik))
		end if
	Next
end if
set rsC=nothing

xmlcat=Replace(oXML.xml,vbcrlf,"")

SavedFile1 = "xmlcat.asp"
findit = Server.MapPath(Savedfile1)
Set fso = server.CreateObject("Scripting.FileSystemObject")
Err.number=0
Set f = fso.CreateTextFile(findit, true)
f.Write "<" & Chr(37) & "xmlcat=""" & Replace(Replace(oXML.xml,"""",""""""),vbcrlf,"") & """" & Chr(37) & ">"
f.close
%>