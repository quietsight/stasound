<%
' ***********************************************************************
' ** Version:			1.0	- 18/01/2002											**
' ** Author:			Mat Peck													**
' ** Function:		Contains simple procedures to encode, encrypt, 	 	**
' **					decode, decrypt and split the information POSTed 	**
' ** 					to and from VSP Form.									**
' **																				**
' ** Revision History:															**
' **	Version	Author			Date and notes								**
' **		1.0		Mat Peck		18/01/2002 - First release					**
' **		1.1		Mat Peck		07/03/2002 - Base64 routines patched		**
' ***********************************************************************


' ** Set variables to indentify the vendor **

	'VendorName="testvendor"
	'Password="testvendor"
	
' ** Your server's IP address or dns name and web app directory.  Full qualified **
' ** Examples : MyServer="https://www.newco.com/ASPFormKit/", MyServer="192.168.0.1/ASPFormKit", MyServer="http://localhost/ASPFormKit/" **

	MyServer="../../pc/"	
	
	eoln = chr(13) & chr(10)

' ** Set up the Base64 arrays

	const BASE_64_MAP_INIT ="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	dim nl
   	dim Base64EncMap(63)
   	dim Base64DecMap(127)
	initcodecs



' ** The initCodecs() routine initialises the Base64 arrays

   PUBLIC SUB initCodecs()
          nl = "<P>" & chr(13) & chr(10)
          dim max, idx
          max = len(BASE_64_MAP_INIT)
          for idx = 0 to max - 1
               Base64EncMap(idx) = mid(BASE_64_MAP_INIT, idx + 1, 1)
          next
          for idx = 0 to max - 1
               Base64DecMap(ASC(Base64EncMap(idx))) = idx
          next
   END SUB

' ** Base 64 Encoding function **

   PUBLIC FUNCTION base64Encode(plain)

			call initCodecs

          if len(plain) = 0 then
               base64Encode = ""
               exit function
          end if

          dim ret, ndx, by3, first, second, third
          by3 = (len(plain) \ 3) * 3
          ndx = 1
          do while ndx <= by3
               first  = asc(mid(plain, ndx+0, 1))
               second = asc(mid(plain, ndx+1, 1))
               third  = asc(mid(plain, ndx+2, 1))
               ret = ret & Base64EncMap(  (first \ 4) AND 63 )
               ret = ret & Base64EncMap( ((first * 16) AND 48) + ((second \ 16) AND 15 ) )
               ret = ret & Base64EncMap( ((second * 4) AND 60) + ((third \ 64) AND 3 ) )
               ret = ret & Base64EncMap( third AND 63)
               ndx = ndx + 3
          loop
          ' check for stragglers
          if by3 < len(plain) then
               first  = asc(mid(plain, ndx+0, 1))
               ret = ret & Base64EncMap(  (first \ 4) AND 63 )
               if (len(plain) MOD 3 ) = 2 then
                    second = asc(mid(plain, ndx+1, 1))
                    ret = ret & Base64EncMap( ((first * 16) AND 48) +((second \16) AND 15 ) )
                    ret = ret & Base64EncMap( ((second * 4) AND 60) )
               else
                    ret = ret & Base64EncMap( (first * 16) AND 48)
                    ret = ret & "="
               end if
               ret = ret & "="
          end if

          base64Encode = ret
     END FUNCTION

' ** Base 64 decoding function **

     PUBLIC FUNCTION base64Decode(scrambled)

          if len(scrambled) = 0 then
               base64Decode = ""
               exit function
          end if

          ' ignore padding
          dim realLen
          realLen = len(scrambled)
          do while mid(scrambled, realLen, 1) = "="
               realLen = realLen - 1
          loop
          do while instr(scrambled," ")<>0
          		scrambled=left(scrambled,instr(scrambled," ")-1) & "+" & mid(scrambled,instr(scrambled," ")+1)
          loop
          dim ret, ndx, by4, first, second, third, fourth
          ret = ""
          by4 = (realLen \ 4) * 4
          ndx = 1
          do while ndx <= by4
               first  = Base64DecMap(asc(mid(scrambled, ndx+0, 1)))
               second = Base64DecMap(asc(mid(scrambled, ndx+1, 1)))
               third  = Base64DecMap(asc(mid(scrambled, ndx+2, 1)))
               fourth = Base64DecMap(asc(mid(scrambled, ndx+3, 1)))
               ret = ret & chr( ((first * 4) AND 255) +   ((second \ 16) AND 3))
               ret = ret & chr( ((second * 16) AND 255) + ((third \ 4) AND 15) )
               ret = ret & chr( ((third * 64) AND 255) +  (fourth AND 63) )
               ndx = ndx + 4
          loop
          ' check for stragglers, will be 2 or 3 characters
          if ndx < realLen then
               first  = Base64DecMap(asc(mid(scrambled, ndx+0, 1)))
               second = Base64DecMap(asc(mid(scrambled, ndx+1, 1)))
               ret = ret & chr( ((first * 4) AND 255) +   ((second \ 16) AND 3))
               if realLen MOD 4 = 3 then
                    third = Base64DecMap(asc(mid(scrambled,ndx+2,1)))
                    ret = ret & chr( ((second * 16) AND 255) + ((third \ 4) AND 15) )
               end if
          end if

          base64Decode = ret
     END FUNCTION


' ** The SimpleXor encryption algorithm. **
' ** NOTE:		This is a placeholder really.  Future releases of VSP Form will use AES or TwoFish.  Proper encryption **
' ** 			This simple function and the Base64 will deter script kiddies and prevent the "View Source" type tampering **
' **			It won't stop a half decent hacker though, but the most they could do is change the amount field to something **
' **			else, so provided the vendor checks the reports and compares amounts, there is no harm done.  It's still **
' **			more secure than the other PSPs who don't both encrypting their forms at all **

Public Function SimpleXor(InString,Key)
    Dim myIN, myKEY, myC, myPub
    Dim Keylist()
    
    myIN = InString
    myKEY = Key
    
    redim KeyList(Len(myKEY))
    
    i = 1
    do while i<=Len(myKEY)
        KeyList(i) = Asc(Mid(myKEY, i, 1))
        i = i + 1
    loop       
    
    j = 1
    i = 1
    do while i<=Len(myIn)
        myC = myC & Chr(Asc(Mid(myIN, i, 1)) Xor KeyList(j))
        i = i + 1
        If j = Len(myKEY) Then j = 0
        j = j + 1
    loop
 
    SimpleXor = myC
End Function


' ** The getToken function. **
' ** NOTE:		A function of convenience that extracts the value from the "name=value&name2=value2..." VSP reply string **
' **			Works even if one of the values is a URL containing the & or = signs.  **

public function getToken(thisString,thisToken)

	' Can't just rely on & characters because these may be provided in the URLs.
	Dim Tokens
	Dim subString
	Tokens = Array("Status","StatusDetail","VendorTxCode","VPSTxID","TxAuthNo","AVSCV2","Amount")
	
	if instr(thisString,thisToken+"=")=0 then
	
		'  If the token isn't present, empty the output.  We can error later
		getToken=""
		
	else
		
		' Right get the rest of the string
		subString=mid(thisString,instr(thisString,thisToken)+len(thisToken)+1)
		
		' Now strip off all remaining tokens if they are present.
		
		i=0
		do while i<7
		
			'Find the next token and lop it off
			if Tokens(i)<>thisToken then
			
				if instr(subString,"&"+Tokens(i))<>0 then 
					substring=left(substring,instr(subString,"&"+Tokens(i))-1)
				end if
						
			end if
			
			i = i +1
		
		loop	
		
		getToken=subString
	
	end if

  
end function



%>