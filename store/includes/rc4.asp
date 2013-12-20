<%  
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

Dim sbox(255)
Dim key(255)


Sub RC4Initialize(strPwd)
	':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':::  This routine called by EnDeCrypt function. Initializes the :::
	':::  sbox and the key array)                                    :::
	':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

	dim tempSwap
	dim a
	dim b

	intLength=len(strPwd)
	For a=0 To 255
		 key(a)=asc(mid(strpwd, (a mod intLength)+1, 1))
		 sbox(a)=a
	next

	b=0
	For a=0 To 255
		 b=(b + sbox(a) + key(a)) Mod 256
		 tempSwap=sbox(a)
		 sbox(a)=sbox(b)
		 sbox(b)=tempSwap
	Next

End Sub
   
Function EnDeCrypt(plaintxt, psw)
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

	RC4Initialize psw
	
	'restore special character
	plaintxt1 = replace(plaintxt,chr(1) & "DD",chr(0))

	For a=1 To Len(plaintxt1)
		 i=(i + 1) Mod 256
		 j=(j + sbox(i)) Mod 256
		 temp=sbox(i)
		 sbox(i)=sbox(j)
		 sbox(j)=temp

		 k=sbox((sbox(i) + sbox(j)) Mod 256)

		 cipherby=Asc(Mid(plaintxt1, a, 1)) Xor k
		 cipher=cipher & Chr(cipherby)
	Next

	EnDeCrypt=cipher
	' remove quotes and special character 
	EnDeCrypt=replace(EnDeCrypt,"'","''")
	EnDeCrypt=replace(EnDeCrypt,chr(0),chr(1) & "DD")      
End Function
%>