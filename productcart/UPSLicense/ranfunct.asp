<%
     Function gen_pass(max_num)

     dim gen_array(62)
' ------- Setup array of characters to chose from ------
    
     gen_array(0) = "0"
     gen_array(1) = "1"
     gen_array(2) = "2"
     gen_array(3) = "3"
     gen_array(4) = "4"
     gen_array(5) = "5"
     gen_array(6) = "6"
     gen_array(7) = "7"
     gen_array(8) = "8"
     gen_array(9) = "9"
     gen_array(10) = "A"
     gen_array(11) = "B"
     gen_array(12) = "C"
     gen_array(13) = "D"
     gen_array(14) = "E"
     gen_array(15) = "F"
     gen_array(16) = "G"
     gen_array(17) = "H"
     gen_array(18) = "I"
     gen_array(19) = "J"
     gen_array(20) = "K"
     gen_array(21) = "L"
     gen_array(22) = "M"
     gen_array(23) = "N"
     gen_array(24) = "O"
     gen_array(25) = "P"
     gen_array(26) = "Q"
     gen_array(27) = "R"
     gen_array(28) = "S"
     gen_array(29) = "T"
     gen_array(30) = "U"
     gen_array(31) = "V"
     gen_array(32) = "W"
     gen_array(33) = "X"
     gen_array(34) = "Y"
     gen_array(35) = "Z"
     gen_array(36) = "a"
     gen_array(37) = "b"
     gen_array(38) = "c"
     gen_array(39) = "d"
     gen_array(40) = "e"
     gen_array(41) = "f"
     gen_array(42) = "g"
     gen_array(43) = "h"
     gen_array(44) = "i"
     gen_array(45) = "j"
     gen_array(46) = "k"
     gen_array(47) = "l"
     gen_array(48) = "m"
     gen_array(49) = "n"
     gen_array(50) = "o"
     gen_array(51) = "p"
     gen_array(52) = "q"
     gen_array(53) = "r"
     gen_array(54) = "s"
     gen_array(55) = "t"
     gen_array(56) = "u"
     gen_array(57) = "v"
     gen_array(58) = "w"
     gen_array(59) = "x"
     gen_array(60) = "y"
     gen_array(61) = "z"

     Randomize
' ------- Generate the string until the length of max_num is met ------
     do while len(output) < max_num
          num = gen_array(Int((61 - 0 + 1) * Rnd + 0))
          output = output + num
     loop

' ------- Let function result = output ------

     gen_pass = output
     End Function %>
