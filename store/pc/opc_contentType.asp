<%
Dim pcv_strCharSet
'////////////////////////////////////////////////////////////////
'// Choose a Character Set
'////////////////////////////////////////////////////////////////
' - ISO-8859-1
' - UTF-8
'////////////////////////////////////////////////////////////////
pcv_strCharSet = "ISO-8859-1"

Public Sub SetContentType
	If ucase(pcv_strCharSet)="UTF-8" Then
		response.ContentType = "text/html"
	Else
		response.ContentType = "text/html"
		Response.CharSet = pcv_strCharSet
	End If
End Sub

Private Function URLDecode(encodedstring)	
	If ucase(pcv_strCharSet)="UTF-8" Then
		URLDecode = encodedstring
	Else
		URLDecode = utf8_decode(encodedstring)
	End If
End Function
%>
<script language="javascript" type="text/javascript" runat="server">
function utf8_decode ( str_data ) {
    // Converts a UTF-8 encoded string to ISO-8859-1  
    // 
    // version: 909.322
    // discuss at: http://phpjs.org/functions/utf8_decode
    // +   original by: Webtoolkit.info (http://www.webtoolkit.info/)
    // +      input by: Aman Gupta
    // +   improved by: Kevin van Zonneveld (http://kevin.vanzonneveld.net)
    // +   improved by: Norman "zEh" Fuchs
    // +   bugfixed by: hitwork
    // +   bugfixed by: Onno Marsman
    // +      input by: Brett Zamir (http://brett-zamir.me)
    // +   bugfixed by: Kevin van Zonneveld (http://kevin.vanzonneveld.net)
    // *     example 1: utf8_decode('Kevin van Zonneveld');
    // *     returns 1: 'Kevin van Zonneveld'
    var tmp_arr = [], i = 0, ac = 0, c1 = 0, c2 = 0, c3 = 0;
    
    str_data += '';
    
    while ( i < str_data.length ) {
        c1 = str_data.charCodeAt(i);
        if (c1 < 128) {
            tmp_arr[ac++] = String.fromCharCode(c1);
            i++;
        } else if ((c1 > 191) && (c1 < 224)) {
            c2 = str_data.charCodeAt(i+1);
            tmp_arr[ac++] = String.fromCharCode(((c1 & 31) << 6) | (c2 & 63));
            i += 2;
        } else {
            c2 = str_data.charCodeAt(i+1);
            c3 = str_data.charCodeAt(i+2);
            tmp_arr[ac++] = String.fromCharCode(((c1 & 15) << 12) | ((c2 & 63) << 6) | (c3 & 63));
            i += 3;
        }
    }

    return tmp_arr.join('');
}
</script>
