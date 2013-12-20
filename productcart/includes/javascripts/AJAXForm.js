// RPC Functions

//Function to intialize an XMLHTTPRequest object (cross-browser)
function createDirectRequestObject () {
	var xmlhttp=false;
	/*@cc_on @*/
	/*@if (@_jscript_version >= 5)
	// JScript gives us Conditional compilation, we can cope with old IE versions.
	// and security blocked creation of the objects.
	try {
		xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
	} catch (e) {
	 	try {
			xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
	  	} catch (E) {
	   		xmlhttp = false;
	  	}
	}
	@end @*/
	
	if (!xmlhttp && typeof XMLHttpRequest!='undefined') {
		xmlhttp = new XMLHttpRequest();
	}
	
	return xmlhttp;
}




//Function to submit a form using xmlHttpRequest object
function submitForm (form, debug, resultFunc, validationFunc) {
	//first we call the validation function if it was present.
	
	//Assume the form is valid
	var validResult = true;
	if (validationFunc != null) {
		try {
			//The validation Function should return false if the form is not valid
			validResult = validationFunc(form);
		} catch (e) {
			if (debug) {
				alert ("Validation function could not be called due to " + e);
			}
		}
	}
	
	if (validResult != true) {
		//halt the process now, form was not valid
		return false;
	}
	
	//Form can now be considered valid, so let's proceed
	
	//First we need to get some information about the form itself
	var url = form.getAttribute("action");
	var enctype = form.getAttribute("enctype");
	var method = form.getAttribute("method");
	
	
	
	//now some quick error checking on the form values
	if (!url || url.length < 1) { 
		//If Developer Debugging is enabled then output a developer message
		if (debug == true) {
			alert ("The form does not have a valid action attribute and so can't be sumbitted.");
		} else {
			//Output a friendly error message for users
			alert ("The form cannot be submitted at this time. Please try again later.");
		}
		
		return false;
	}
	
	
	
	if (!enctype || enctype.length < 1) {
		//Define default value
		enctype = "application/x-www-form-urlencoded";
	}
	
	if (method.toUpperCase() != "POST" && method.toUpperCase() != "GET") {
		if (debug == true) {
			alert ("The method for the form was defined as '" + method +"'. Only POST and GET are valid; assuming POST...");
		}
		method = "POST";
	}
	
	//Next we use the encapsulation type to try and determine how to format the data
	switch (enctype) {
		case ("application/x-www-form-urlencoded"):
			//Standard URL encoding
			
			var dataString = "";
			
			//Can't just use the form.elements array because fieldsets appear as elements...
			var theseFields = getFormFieldsInEle(form);
			
			//Loop over all the form elements to build the data string
			for (var i=0; i < theseFields.length; i++) {
				//Flag to determine if a value was added to the dataString this time around or not...
				//used later to decide whether to add an & or not
				var updatedVal = false;
				
				
				//We can't submit values for unidentified fields
				if (theseFields[i].name != "") {
					
					
					//Special case for radio buttons
					var isRadio = false;
					if (theseFields[i].type) {
						if (theseFields[i].type.toUpperCase() == "RADIO") {
							isRadio = true;
						}
					}
					if (isRadio) {
						if (dataString.indexOf(escape(theseFields[i].name) + "=") < 0) {
						
							var radioValue = getRadioValue(form, theseFields[i].name);
							
							if (radioValue != "") {
								dataString += escape(String(theseFields[i].name).replace(/ /, "+")) + "=" 
						
						
								dataString += escape(String(radioValue).replace(/ /, "+"));
							
								updatedVal = true;
							}
							
						}
					} else {
						//Need to make an exception to handle select multiple field types...
						if (theseFields[i].multiple) {
							//Exception to the rule for multiple select form fields
							var thisValueArr = findSelectValue(theseFields[i]);
							
							for (var p =0; p < thisValueArr.length; p++) {
								dataString += escape(String(theseFields[i].name).replace(/ /, "+")) + "=" + escape(String(thisValueArr[p].value).replace(/ /, "+"));
								if (p < thisValueArr.length -1) {
									dataString += "&";
								}
								updatedVal = true;
							}
						} else {
					
							//All other form field types work this way...
							dataString += escape(String(theseFields[i].name).replace(/ /, "+")) + "=" + escape(String(theseFields[i].value).replace(/ /, "+"));
							updatedVal = true;
						}
					}
					
				}
				
				if (updatedVal) {
					dataString += "&";
				}
				
				
				
				
			}
			
			//the last thing to do is strip off an extra & if it occurs
			if (dataString.lastIndexOf("&") == dataString.length - 1) {
				dataString = dataString.substring(0, dataString.length -1);
			}
			
			
			
			break;
			
		case ("text/xml"):
			//This is meant to be a stepping stone towards XForms. Using this method allows you to build server-side processing
			//that accepts form data as an XML structure rather than a simple name:value pair. The structure is based on the 
			//semantic markup of the XHTML form in the current document.
			
			//Start by trying to break the form up into fieldsets
			var fieldSets = form.getElementsByTagName("fieldset");
			
			if (fieldSets.length < 1) {
				//No field sets defined so just use the form instead
				fieldSets = new Array(form);
			}
			
			//Start the XML document with the XML header
			var dataString = '<?xml version="1.0" encoding="iso-8859-1"?>';
			
			//Add the root element
			dataString += "<form>";
			
			//As prep work we'll use an outside function to get all of the "label" elements within the form and 
			//put them into an associative array indexed by the fields that they relate to; this way when we loop over
			//the fields we can quickly and easily grab the appropriate labels.
			var formLabelsArr = getFormLabels(form)
			
			//Loop over and output each fieldset
			for (var i=0; i < fieldSets.length; i++) {
				//Try and get a label for the field set
				var fieldSetLabel = fieldSets[i].getElementsByTagName("legend");
				
				if (fieldSetLabel.length > 0) {
					fieldSetLabel = fieldSetLabel.item(0);
				} else {
					fieldSetLabel = false;
				}
				
				//Output the opening tag (with label attrib if present)
				dataString += "<fieldset";
				if (fieldSetLabel) {
					dataString += " label=\"" + getTextOfNode(fieldSetLabel) + "\"";
				}
				dataString += ">";
				
				//Now loop over each element in the fieldset
				var theseFields = getFormFieldsInEle(fieldSets[i]);
				
				
				for (var p=0; p < theseFields.length; p++) {
					
					var thisField = theseFields[p];
					
					//Info about this field
					var thisFieldInfo = new Array();
					
					//The field value
					var thisFieldValue = "";
					
					
					if (thisField.name) {
						thisFieldInfo["name"] = thisField.name;
					}
					
					if (thisField.id) {
						thisFieldInfo["id"] = thisField.id;
					}
					
					//Get the info about this field
					switch (thisField.tagName.toUpperCase()) {
						case ("INPUT"):
							thisFieldInfo["type"] = thisField.type;
							
							//special case for radio buttons
							if (thisField.type.toUpperCase() == "RADIO") {
								if (!(dataString.indexOf("name=\"" + thisField.name + "\"") > 0 || dataString.indexOf("id=\"" + thisField.id + "\"") > 0)) {
						
									thisFieldValue = getRadioValue(form, thisField.name);
								}
							} else {
								thisFieldValue = thisField.value;
							}
							break;
							
						case ("TEXTAREA"):
							thisFieldInfo["type"] = "textarea";
							thisFieldValue = thisField.value;
							break;
						case ("SELECT"):
							thisFieldInfo["type"] = "select";
							thisFieldInfo["allowMultiple"] = thisField.multiple;
							
							thisFieldValue = findSelectValue(thisField);
							
							break;
					}
										
					
					
					//Try to find a label for this field
					var labelEle = null;
					if (thisField.id) {
						if (formLabelsArr[thisField.id]) {
							labelEle = formLabelsArr[thisField.id];
						}
					} else if (thisField.name) {
						if (formLabelsArr[thisField.name]) {
							labelEle = formLabelsArr[thisField.name];
						}
					}
					
					if (labelEle != null) {
						thisFieldInfo["label"] = getTextOfNode(labelEle);
					}
					
					if (thisFieldValue != "") {
						//Now output the fields with info and values
						dataString += "<field ";
						
						for (var t in thisFieldInfo) {
							dataString += t + "=\"" + thisFieldInfo[t] + "\" ";
						}
						
						dataString += ">";
						
						
						if (String(typeof(thisFieldValue)).toLowerCase() == "object") {
	
							//A select field with multiple values 
							for (var j = 0; j < thisFieldValue.length; j++) {
								dataString += "<value";
								if (thisFieldValue[j].getAttribute("optGroup")) {
									dataString += " optGroup=\"" + thisFieldValue[j].getAttribute("optGroup") + "\" ";
								}
								dataString += ">" + thisFieldValue[j].value + "</value>";
							}
						} else {
							dataString += "<value>" + thisFieldValue + "</value>";
						}
						
						dataString += "</field>";
					}
					
				}

							
					
					
						
					
				
				
				
				
				
				//close this fieldSet
				dataString += "</fieldset>";
			
			}
			
			//close the root element
			dataString += "</form>";
			
			
			break;
			
		default:
			if (debug) {
				alert("The form is set to an encapsulation type of:" + enctype + " \n\
This type of encapsulation is not supported by this script. Please use either \n\
application/x-www-form-urlencoded or \n\
text/xml");

			}
			break;
			
	
	}
	
	
	if (dataString) {
		//Submit the form
			
		//Create the request object
		var xmlHttpObj = createDirectRequestObject();
		
		//Create the callback
		xmlHttpObj.onreadystatechange = function () {
			if (xmlHttpObj.readyState == 4) {
				if (resultFunc != null) {
					try {
						resultFunc(xmlHttpObj.responseText, xmlHttpObj.responseXML);
					} catch (e) {
						if (debug) {
							alert ("Response function could not be called:" + e);
						}
					}
				}
			}
		}
		
		//Begin the transaction
		xmlHttpObj.open(method, url, true);
		xmlHttpObj.setRequestHeader("Content-Type", enctype);
		xmlHttpObj.send(dataString);
		
		
		
	}
		
		 
	return false;
}


//Function to get the value of a radio group
function getRadioValue (form, radioGroupName) {
	for (var i=0; i < form.elements.length; i++) {
		if (form.elements[i].name == radioGroupName) {
			if (form.elements[i].checked) {
				return form.elements[i].value;
			}
		}
	}
	return "";
}


//function to get all form fields (of all types) that are contained by a given element
function getFormFieldsInEle(element, onlyThisType) {
	//quick error check
	if (onlyThisType != null) {
		if (onlyThisType.toLowerCase() != "input" && onlyThisType.toLowerCase() != "textarea" && onlyThisType.toLowerCase() != "select") {
			onlyThisType = null;
		}
	}
	
	
	if (onlyThisType == null) {
		var typesArr = new Array("input", "textarea", "select");
	}
	
	//Array of all the returned elements
	var resultsArr = new Array();
	
	if (onlyThisType == null) {
		//Loop over and grab all of them
		for (var p=0; p < typesArr.length; p++) {
			var theseFields = element.getElementsByTagName(typesArr[p]);
			
			if (theseFields.length > 0) {
			
				for (var j =0; j < theseFields.length; j++) {
					resultsArr.push(theseFields[j]);
				}
				
			}
		}
	} else {
		//Just the fields of the given type 
		resultsArr = element.getElementsByTagName(typesArr[p]);
	}
	
	return resultsArr;
		
	
}


//function to get all the labels of a given form and put them into an associative array indexed by the fields that they relate to.
function getFormLabels(form) {
	var labels = form.getElementsByTagName("label");
	
	//This will be our associative array
	var resultArr = new Array();
	
	for (var j =0; j < labels.length; j++) {
		var thisLabel = labels[j];
		
		//First check to see if the "for" attribute has been set, and if so use that as the index;
		if (thisLabel.getAttribute("for")) {
			resultArr[thisLabel.getAttribute("for")] = thisLabel;
		} else {
			//Now we have to do it the hard way and find out the id of the field that this label contains
			//We do this using our getFormFieldsInEle functoin
			var labelFields = getFormFieldsInEle(thisLabel);
			
			if (labelFields.length > 0) {
				//there should only ever be one form element within a label, so we can assume it's the first returned item
				var thisField = labelFields[0];
				
				if (thisField.id) {
					resultArr[thisField.id] = thisLabel;
				} else if (thisField.name) {
					resultArr[thisField.name] = thisLabel;
				}
			}
		}
	}
	
	return resultArr;

}

//Function to get the text contents of a node (but not nested nodes)
//this function is used to retrieve the text value of labels, just incase the
//label element wraps either side of the form field.
function getTextOfNode(node) {
	var textContent = "";
	if (node.childNodes.length > 0) {
		var children = node.childNodes;
		
		for (var q=0; q < children.length; q++) {
			if (children[q].nodeType == 3) {
				textContent += children[q].nodeValue;
			}
		}
	}
	
	//strip out new lines
	var reg = /\n/
	var reg2 = /\r/
	textContent = textContent.replace(reg, "");
	textContent = textContent.replace(reg2, "");
	
	return textContent;
}


//Function to find the value of a select element. (this is needed to allow for multiple selections;
//it also provides opt group information if present

function findSelectValue (field) {
	//first thing to do is find the value of 
	
	var results = new Array();
	
	//Loop over each option, see if it's selected, and get it's parent if it is
	for (var q=0; q < field.options.length; q++) {
		var thisOption = field.options[q];
		var thisOptGroup = null;
		
		if (thisOption.selected) {
			if (thisOption.parentNode.tagName.toUpperCase() == "OPTGROUP") {
				if (thisOption.parentNode.getAttribute("label")) {
					thisOptGroup = thisOption.parentNode.getAttribute("label");
				}
			}
			
			if (thisOptGroup) {
				thisOption.setAttribute("optGroup", thisOptGroup);
			}
			results.push(thisOption);
			
		}
	}
	
	return results;
	
}
