<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Function RndTo(intNum, intRn)
	RndTo= Int((intNum / intRn)+.5) * intRn
End Function 

function pc_dimensions(number)
	if number="" OR number=0 then
		number=0
		pc_dimensions=0
	end if
	' round number
	if number<>0 then
		'pc_dimensions=round(number,2)
		pc_dimensions=RndTo(number,.01)
	end if
	
	' locate dec division
	dim indexPoint
	indexPoint=instr(pc_dimensions,".")
	
	' for integer, add .00
	if indexPoint=0 then
			pc_dimensions=Cstr(pc_dimensions)+".00"
	end if

	' calculate if 0 or 00
	dim pc_dimensionsLarge, decPart
	pc_dimensionsLarge=len(pc_dimensions)
	decPart=right(pc_dimensions,pc_dimensionsLarge-indexPoint)
	
	' add to original numbers
	dim g
	for g=0 to (1- (pc_dimensionsLarge-indexPoint))
			pc_dimensions=Cstr(pc_dimensions)+"0"    
	next
	
end function

%>