<%
Session.CodePage = 1252
     '|----------------------------------------------------|
     '| Encryption / Decryption                            |
     '| version 3.0  |  added ASP functionallity           |
     '| By Willem Bult                                     |
     '| www.willem.nl.nu                                   |
     '|                                                    |
     '| Barry L Beattie made the first step when he added  |
     '| ASP functionallity to the form in version 2        |
     '|                                                    |
     '| In version 3.0 everything is in ASP                |
     '| This means that the source will be hidden from the |
     '| client, which means extra security                 |
     '|                                                    |
     '| Also new in this version: a swap function          |
     '|                                                    |
     '| Free to use as long as this message remains intact |
     '|----------------------------------------------------|


  Function Encrypt(Uncoded, Password)
    Uncoded=Swap(Uncoded)                            'Run the text through the swap function
    
    For Char=1 to LEN(Uncoded)
      TxtChar=ASC(MID(Uncoded, Char, 1))             'Store character codes of text and password
      PwdChar=ASC(MID(Password, (Char MOD LEN(Password) + 1), 1))

      NewChar=TxtChar + PwdChar                      'Combine them into one new character code
      If NewChar > 255 Then NewChar=NewChar - 255    'Charactercode can't be >255 or <1 
      
      Encrypt=Encrypt & Chr(NewChar)                 'Add new charactercode
    Next
    
    Uncoded=Swap(Uncoded)                            'back-swap the text so it will be displayed correctly (not necessary for successful encryption)
  End Function
  
  Function Decrypt(Coded, Password)
    
    For Char=1 to LEN(Coded)
      CodChar=ASC(MID(Coded, Char, 1))               'Store character codes of text and password
      PwdChar=ASC(MID(Password, (Char MOD LEN(Password) + 1), 1))
      
      NewChar=CodChar - PwdChar                      'Restore the original charactercode
      if NewChar < 1 then NewChar=NewChar + 255
      
      Decrypt=Decrypt & Chr(NewChar)                 'Add original charactercode
    Next
    
    Decrypt=Swap(Decrypt)                            'Swap the result to get the original string
  End Function
  
  Function Swap(Inp)
    Dim InpTemp(3)                                     'Make array with 4 positions
    
    For Char=1 to LEN(Inp) step 4                    'Walk through the string
	  
	  if Char + 2 < LEN(Inp) then                      'If there are enough characters left to do a swap
	    
	    for i=0 to 3                                 
	      InpTemp(i)=MID(Inp, Char + i, 1)           'Store 4 characters in array
        next
        
        if LEN(Inp) MOD 4 > 1 then                     'I used two ways of swapping (for extra security), it depends on the length of the string wich swap will be done
          Outp=Outp & InpTemp(2) & InpTemp(3) & InpTemp(0) & InpTemp(1)
        else
		  Outp=Outp & InpTemp(3) & InpTemp(2) & InpTemp(1) & InpTemp(0)
	    end if
	  
	  else
	    Outp=Outp & MID(Inp, Char, LEN(Inp) - Char + 1)   'If swap couldn't be made, just add the remaining characters
	  end if
    
    Next
    
    Swap=Outp                                        'Return the swapped string
  End Function
  
  %>
