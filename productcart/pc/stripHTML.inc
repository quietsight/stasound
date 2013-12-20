<%

	'[ClearHTMLTags]
	
	'Coded by Jóhann Haukur Gunnarsson
	'joi@innn.is
	
	'	Purpose: This function clears all HTML tags from a string using Regular Expressions.
	'	 Inputs: strHTML;	A string to be cleared of HTML TAGS
	'		 intWorkFlow;	An integer that if equals to 0 runs only the RegExp filter
	'							  .. 1 runs only the HTML source render filter
	'							  .. 2 runs both the RegExp and the HTML source render
	'							  .. >2 defaults to 0
	'	Returns: A string that has been filtered by the function
	
	
	function ClearHTMLTags(strHTML, intWorkFlow)
	
		'Variables used in the function
		
		dim regEx, strTagLess
		
		'---------------------------------------
		strTagless = strHTML
		'Move the string into a private variable
		'within the function
		'---------------------------------------

		'regEx initialization

		'---------------------------------------
		set regEx = New RegExp 
		'Creates a regexp object		
		regEx.IgnoreCase = True
		'Don't give frat about case sensitivity
		regEx.Global = True
		'Global applicability
		'---------------------------------------


		'Phase I
		'	"bye bye html tags"


		if intWorkFlow <> 1 then
		
			'---------------------------------------
			regEx.Pattern = "<[^>]*>"
			'this pattern mathces any html tag
			strTagLess = regEx.Replace(strTagLess, "")
			'all html tags are stripped
			'---------------------------------------
						
		end if
		
		
		'Phase II
		'	"bye bye rouge leftovers"
		'	"or, I want to render the source"
		'	"as html."

		'---------------------------------------
		'We *might* still have rouge < and > 
		'let's be positive that those that remain
		'are changed into html characters
		'---------------------------------------	
		

		if intWorkFlow > 0 and intWorkFlow < 3 then
		
		
			regEx.Pattern = "[<]"
			'matches a single <
			strTagLess = regEx.Replace(strTagLess, "&lt;")

			regEx.Pattern = "[>]"
			'matches a single >
			strTagLess = regEx.Replace(strTagLess, "&gt;")
			'---------------------------------------
			
		end if
		
		
		'Clean up
							
		'---------------------------------------
		set regEx = nothing
		'Destroys the regExp object
		'---------------------------------------	
		
		'---------------------------------------
		ClearHTMLTags = strTagLess
		'The results are passed back
		'---------------------------------------
		
	end function

%>


