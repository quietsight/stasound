<%
' Gen Key Functions
Function gen_pass(GEN_NUM)

	dim gen_array(26)
	' ------- Setup array of characters to chose from ------
	
	gen_array(0) = "A"
	gen_array(1) = "B"
	gen_array(2) = "C"
	gen_array(3) = "D"
	gen_array(4) = "E"
	gen_array(5) = "F"
	gen_array(6) = "G"
	gen_array(7) = "H"
	gen_array(8) = "I"
	gen_array(9) = "J"
	gen_array(10) = "K"
	gen_array(11) = "L"
	gen_array(12) = "M"
	gen_array(13) = "N"
	gen_array(14) = "O"
	gen_array(15) = "P"
	gen_array(16) = "Q"
	gen_array(17) = "R"
	gen_array(18) = "S"
	gen_array(19) = "T"
	gen_array(20) = "U"
	gen_array(21) = "V"
	gen_array(22) = "W"
	gen_array(23) = "X"
	gen_array(24) = "Y"
	gen_array(25) = "Z"
	
	Randomize
	' ------- Generate the string until the length of max_num is met ------
	do while len(output) < GEN_NUM
		num = gen_array(Int((25 - 0 + 1) * Rnd + 0))
		output = output + num
	loop
	
	' ------- Let function result = output ------
	
	gen_pass = output
End Function
		 
Function gen2_pass(GEN_NUM)

	dim gen2_array(10)
	' ------- Setup array of characters to chose from ------
	
	gen2_array(0) = "0"
	gen2_array(1) = "1"
	gen2_array(2) = "2"
	gen2_array(3) = "3"
	gen2_array(4) = "4"
	gen2_array(5) = "5"
	gen2_array(6) = "6"
	gen2_array(7) = "7"
	gen2_array(8) = "8"
	gen2_array(9) = "9"
	
	Randomize
	' ------- Generate the string until the length of max_num is met ------
	do while len(output) < GEN_NUM
		num = gen2_array(Int((9 - 0 + 1) * Rnd + 0))
		output = output + num
	loop
	' ------- Let function result = output ------
	
	gen2_pass = output
End Function
		 		 
%>