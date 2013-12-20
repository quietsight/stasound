<%  
Session.LCID=1033
Session.Codepage=1252

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::                                                             :::
':::  This script performs 'RC4' Stream Encryption               :::
':::  (Based on what is widely thought to be RSA's RC4           :::
':::  algorithm. It produces output streams that are identical   :::
':::  to the commercial products)                                :::
':::                                                             :::
':::  This script is Copyright C 1999 by Mike Shaffer            :::
':::  ALL RIGHTS RESERVED WORLDWIDE                              :::
':::                                                             :::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Dim gbox(255)
Dim gkey(255)


Sub GRC4Initialize(strPwd)
	':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':::  This routine called by GDeCrypt function. Initializes the :::
	':::  gbox and the gkey array)                                    :::
	':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

	dim tempSwap
	dim a
	dim b

	intLength=len(strPwd)
	For a=0 To 255
		 gkey(a)=asc(mid(strpwd, (a mod intLength)+1, 1))
		 gbox(a)=a
	next

	b=0
	For a=0 To 255
		 b=(b + gbox(a) + gkey(a)) Mod 256
		 tempSwap=gbox(a)
		 gbox(a)=gbox(b)
		 gbox(b)=tempSwap
	Next

End Sub
   
Function GDeCrypt(plaintxt, psw)
 ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
 ':::  This routine does all the work. Call it both to ENcrypt    :::
 ':::  and to DEcrypt your data.                                  :::
 ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

	dim temp
	dim a
	dim i
	dim j
	dim k
	dim cipherby
	dim cipher
	dim plaintxt1

	i=0
	j=0

	GRC4Initialize psw
	
	' Restore Problem Chars
	plaintxt1 = replace(plaintxt,chr(1) & "DD",chr(0))
	plaintxt1 = replace(plaintxt1,"RETURN13",chr(13))
	plaintxt1 = replace(plaintxt1,"RETURN10",chr(10))
	plaintxt1 = replace(plaintxt1,"QUOTE",chr(34))

	For a=1 To Len(plaintxt1)
		 i=(i + 1) Mod 256
		 j=(j + gbox(i)) Mod 256
		 temp=gbox(i)
		 gbox(i)=gbox(j)
		 gbox(j)=temp

		 k=gbox((gbox(i) + gbox(j)) Mod 256)

		 cipherby=Asc(Mid(plaintxt1, a, 1)) Xor k
		 cipher=cipher & Chr(cipherby)
	Next

	GDeCrypt=cipher
	' Replace Problem Chars
	GDeCrypt=replace(GDeCrypt,chr(0),chr(1) & "DD") 
	GDeCrypt=replace(GDeCrypt,chr(13),"RETURN13")  
	GDeCrypt=replace(GDeCrypt,chr(10),"RETURN10")
	GDeCrypt=replace(GDeCrypt,chr(34),"QUOTE")
End Function
%>