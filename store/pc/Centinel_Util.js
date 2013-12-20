
function randomEbatesSid() {
    
    var sid = "ebs";
    for(var i = 0; i < 10; i++) {
        var digit =  new String( Math.floor(Math.random() * 10 ) );
        sid = sid + digit;
    } // end for

    return sid + "sbe";

} // end randomEbatesSid

function randomOrderNumber() {
    var orderNumber =  new String( Math.floor(Math.random() * (9999999999 + 1)  ) );
    var pad = ""; 
    for(var i = 0; i < 10 - orderNumber.length; i++) {
        pad += "0";
    }

    return pad + orderNumber;

} // end randomOrderId

function randomAmount() {
    return new String( Math.floor(Math.random() * (9999 + 1)  ) );
} // end randomAmount

/**
 * Adds <option> tags to the ccStart form date picker and selects
 * current year + 1 by default.
 */
function build3DSDateOptions() {

    var d = new Date();
    var yyyy = d.getFullYear();
    for(var i = yyyy; i < yyyy + 10; i++) {
        var o = document.createElement("option");
        o.value = i;
        o.appendChild(document.createTextNode(i));
        if((yyyy + 1) == i) {
            o.selected = "selected";
        }
        document.forms["frm"].expr_yyyy.appendChild(o);

    }
}

function build3DSRecurringEnd() {

    var d = new Date();
    var yyyy = d.getFullYear();
	var nextYear = yyyy + 1;

	document.frm.recurringEnd.value = nextYear + "0101";

}

function singleSubmit(formName) {

    return function() {
        try {
            var e = document.forms[formName].elements;
            for(var i = 0; i < e.length; i++) {
                if(e[i].type == "submit") {
                    e[i].disabled = 'disabled';
                }
            }
        } catch(e){}
        
        return true;
    }

} // end singleSubmit


/**
 * Enables state codes if "US" is the selected country code.
 */
function countryCodeMonitor() {

    var billingCountry = document.forms[0].b_country_code.value;
    var billingState = document.forms[0].b_state;
    if("US" == billingCountry) {
        document.getElementById("b_state_field").innerHTML = getStateSelectionHTML("b_state");
    } else {
        document.getElementById("b_state_field").innerHTML = getStateInputHTML("b_state");
    }

    var shippingCountry = document.forms[0].s_country_code.value;
    var shippingState = document.forms[0].s_state;
    if("US" == shippingCountry) {
        document.getElementById("s_state_field").innerHTML = getStateSelectionHTML("s_state");
    } else {
        document.getElementById("s_state_field").innerHTML = getStateInputHTML("s_state");
    }


} // end countryCodeMonitor


function getStateSelectionHTML(name) {

    return "<select name=\"" + name + "\">" +
            "<option value=\"\">(No selection)</option>" +
            "<option value=\"AK\">AK - Alaska</option>" +
            "<option value=\"AL\">AL - Alabama</option>" +
            "<option value=\"AR\">AR - Arkansas</option>" +
            "<option value=\"AZ\">AZ - Arizona</option>" +
            "<option value=\"CA\">CA - California</option>" +
            "<option value=\"CO\">CO - Colorado</option>" +
            "<option value=\"CT\">CT - Connecticut</option>" +
            "<option value=\"DC\">DC - District Of Columbia</option>" +
            "<option value=\"DE\">DE - Delaware</option>" +
            "<option value=\"FL\">FL - Florida</option>" +
            "<option value=\"GA\">GA - Georgia</option>" +
            "<option value=\"HI\">HI - Hawaii</option>" +
            "<option value=\"IA\">IA - Iowa</option>" +
            "<option value=\"ID\">ID - Idaho</option>" +
            "<option value=\"IL\">IL - Illinois</option>" +
            "<option value=\"IN\">IN - Indiana</option>" +
            "<option value=\"KS\">KS - Kansas</option>" +
            "<option value=\"KY\">KY - Kentucky</option>" +
            "<option value=\"LA\">LA - Louisiana</option>" +
            "<option value=\"MA\">MA - Massachusetts</option>" +
            "<option value=\"MD\">MD - Maryland</option>" +
            "<option value=\"ME\">ME - Maine</option>" +
            "<option value=\"MI\">MI - Michigan</option>" +
            "<option value=\"MN\">MN - Minnesota</option>" +
            "<option value=\"MO\">MO - Missouri</option>" +
            "<option value=\"MS\">MS - Mississippi</option>" +
            "<option value=\"MT\">MT - Montana</option>" +
            "<option value=\"NC\">NC - North Carolina</option>" +
            "<option value=\"ND\">ND - North Dakota</option>" +
            "<option value=\"NE\">NE - Nebraska</option>" +
            "<option value=\"NH\">NH - New Hampshire</option>" +
            "<option value=\"NJ\">NJ - New Jersey</option>" +
            "<option value=\"NM\">NM - New Mexico</option>" +
            "<option value=\"NV\">NV - Nevada</option>" +
            "<option value=\"NY\">NY - New York</option>" +
            "<option value=\"OH\" selected>OH - Ohio</option>" +
            "<option value=\"OK\">OK - Oklahoma</option>" +
            "<option value=\"OR\">OR - Oregon</option>" +
            "<option value=\"PA\">PA - Pennsylvania</option>" +
            "<option value=\"RI\">RI - Rhode Island</option>" +
            "<option value=\"SC\">SC - South Carolina</option>" +
            "<option value=\"SD\">SD - South Dakota</option>" +
            "<option value=\"TN\">TN - Tennessee</option>" +
            "<option value=\"TX\">TX - Texas</option>" +
            "<option value=\"UT\">UT - Utah</option>" +
            "<option value=\"VA\">VA - Virginia</option>" +
            "<option value=\"VT\">VT - Vermont</option>" +
            "<option value=\"WA\">WA - Washington</option>" +
            "<option value=\"WI\">WI - Wisconsin</option>" +
            "<option value=\"WV\">WV - West Virginia</option>" +
            "<option value=\"WY\">WY - Wyoming</option>" +
        "</select>";
}

function getStateInputHTML(name) {
    return "<input type=\"text\" name=\"" + name + "\" value=\"\">";

}
