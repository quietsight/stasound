<!--#include file="encrypt.asp"-->
<%  dim vErr, fieldName, fieldValue, pFieldValue 

const errImg="<img src=../pc/images/pc_required.gif>" 

set vErr=server.createObject("scripting.dictionary") 

sub ckValF 
	for each field in request.form 
		if left(field, 1)="_" then 
			' is validation field , obtain field name
			fieldName=right( field, len( field ) - 1) 
			' obtain field value
			fieldValue=request.form(field) 

			select case lCase(fieldValue) 

				case "required" 
					if trim(request.form(fieldName))="" then     
						vErr(fieldName)=dictLanguage.Item(Session("language")&"_validateform_1") 
					end if

				case "date" 
					if Not isDate(request.form(fieldName)) then 
						vErr(fieldName)=dictLanguage.Item(Session("language")&"_validateform_2")
					end if  
	
				case "number"    
					pFieldValue=request.form(fieldName)   
					if Not isNumeric(pFieldValue) or (instr(pFieldValue,",")<>0) then 
						vErr(fieldName)=dictLanguage.Item(Session("language")&"_validateform_3") & "<i>" & fieldName & "</i>" 
					end if 

				case "intnumber"    
					pFieldValue=request.form(fieldName)   
					if Not isNumeric(pFieldValue) or (instr(pFieldValue,",")<>0) or (instr(pFieldValue,".")<>0) then 
						vErr(fieldName)=dictLanguage.Item(Session("language")&"_validateform_7") & "<i>" & fieldName & "</i>" 
					end if 
				
				case "positiveNumber" 
					if Not isNumeric(request.form(fieldName)) or (fieldName<0) then 
						vErr(fieldName)=dictLanguage.Item(Session("language")&"_validateform_3") & "<i>" & fieldName & "</i>" 
					end if 

				case "email" 
					pFieldValue=request.form(fieldName)
					pFieldValue=replace(pFieldValue," ","")
					if instr(pFieldValue,"@")=0 or instr(pFieldValue,".")=0 then 
						vErr(fieldValue)=dictLanguage.Item(Session("language")&"_validateform_4")
					end if
					if session("email")<>"" AND pFieldValue<>"" then
						if ucase(trim(session("email")))<>ucase(pFieldValue) then
							'check to see if new email exists already in database
							query="SELECT email FROM customers WHERE email='"&trim(pFieldValue)&"';"
							Set conn=Server.CreateObject("ADODB.Connection")
							conn.Open scDSN  
							set rsValObj=server.CreateObject("ADODB.RecordSet")
							set rsValObj=conn.execute(query)
							if NOT rsValObj.eof then
								'email already exists, alert customer that they can't use that email
								vErr(fieldName)=dictLanguage.Item(Session("language")&"_validateform_10")
							end if
							set rsValObj=nothing
							conn.Close
							set conn=nothing
						end if 
					end if
				case "phone" 
					pFieldValue=request.form(fieldName)
					pFieldValue=replace(pFieldValue," ","")
					pFieldValue=replace(pFieldValue,"-","")
					pFieldValue=replace(pFieldValue,".","")
					pFieldValue=replace(pFieldValue,"(","")
					pFieldValue=replace(pFieldValue,")","")
					if Not isNumeric(pFieldValue)  then 
						vErr(fieldName)=dictLanguage.Item(Session("language")&"_validateform_6")
					end if 
			end select 
		end if 
	next 
end sub 

sub validateForm(byVal successPage) 
	if request.ServerVariables("CONTENT_LENGTH") > 0 then 
		ckValF 
		' if no errors, then successPage 
		if vErr.Count=0 then 
			' build success querystring
			tString=Cstr("")
			for each field in request.form   
				if left(field, 1) <> "_" then 
					fieldName=field
					fieldValue=request.form(fieldName)
					If fieldName="password" AND successPage="login.asp" then
						fieldValue=encrypt(fieldValue, 9286803311968)
					End If 
					tString=tString &fieldName& "=" &Server.Urlencode(fieldValue)& "&"
				end if   
			next
			response.redirect successPage&"?"& tString
		end if 
	end if 
end sub 

sub validateFormDb(byVal successPage) 
	if request.ServerVariables("CONTENT_LENGTH") > 0 then 
		ckValF 
		' if no errors, then successPage 
		if vErr.Count=0 then 
			' build data string to save in DB
			updString=Cstr("")
			for each field in request.form   
				' if its not a validation field
				if left(field, 1) <> "_" then 
				fieldName=field
				if fieldName="Email" then fieldName="email"
					fieldValue=request.form(fieldName)     
					updString=updString &fieldName& "=" & fieldValue & "&-"
				end if   
			next
			' save new data into sessionData 
			updString=replace(updString,"'","''")
			updString=replace(updString,"''''","''")
			updString=replace(updString,"""","&quot;")
			query="UPDATE dbSession SET sessionData='" &updString&"' WHERE idDbSession=" &pIdDbSession& " AND randomKey=" &pRandomKey
			Set conn=Server.CreateObject("ADODB.Connection")
			conn.Open scDSN  
			set rsValObj=server.CreateObject("ADODB.RecordSet")
			set rsValObj=conn.execute(query)
			set rsValObj=nothing
			conn.Close
			set conn=nothing
			response.redirect successPage&"?idDbSession="& pIdDbSession& "&randomKey=" &pRandomKey
		end if 
	end if 
end sub 

sub validateError  
	dim countRow
	countRow=cInt(0)
	for each field in vErr 
		if countRow=0 then
			response.write ""
		end if
		response.write "<div class=""pcCPmessage"">" & vErr(field)&"</div>"
		countRow=countRow+1
	next 
	if countRow>0 then
		response.write "<br>"
	end if
end sub 


sub validate( byVal fieldName, byVal validType ) %>
	<input name="_<%=fieldName%>" type="hidden" value="<%=validType%>"> 
	<% if vErr.Exists(fieldName) then 
		response.write errImg 
	end if 
end sub 

sub textbox(byVal fieldName , byVal fieldValue, byVal fieldSize, byVal fieldType) 
	dim lastValue 
	lastValue=request.form(fieldName) 
	select case fieldType
		case "textbox" %>
			<input name="<%=fieldName%>" size="<%=fieldSize%>" value="<%if trim(fieldValue)<>"" then%><%=Server.HTMLEncode(fieldValue)%><%else%><%=Server.HTMLEncode(lastValue)%><%end if%>"> 
			<%
		case "password" %>
			<input name="<%=fieldName%>" type="password" size="<%=fieldSize%>" value="<%if trim(fieldValue)<>"" then%><%=Server.HTMLEncode(fieldValue)%><%else%><%=server.HTMLEncode(lastValue)%><%end if%>"> 
		<%
		case "textarea" %>
			<textarea name="<%=fieldName%>" rows="5" cols="<%=fieldSize%>"><%
			if trim(fieldValue)<>"" then
				response.write fieldValue    
			else
				response.write Server.HTMLEncode(trim(lastValue))
			end if
		%></textarea>
	<% end select 
end sub %>