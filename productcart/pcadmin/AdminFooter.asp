<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, all of its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit http://www.productcart.com.
%>
        </div>
        <div id="pcCPmainRight">
        
        <%
		if pcInt_ShowOrderLegend = 1 then
		%>
        	<div class="pcCPsearchBox" id="cpBox5">
                <div class="CollapsiblePanel" id="cp5">
                  <div class="CollapsiblePanelTab"><h1>Order Status Legend</h1></div>
                    <ul>
                    	<li><img src="images/bluedot.gif"> Pending</li>
                        <li><img src="images/yellowdot.gif"> Processed</li>
                        <li><img src="images/7dot.gif"> Partially Shipped</li>
                        <li><img src="images/8dot.gif"> Shipping</li>
                        <li><img src="images/greendot.gif"> Shipped</li>
                        <li><img src="images/9dot.gif"> Partially Returned</li>
                        <li><img src="images/orangedot.gif"> Returned</li>
                        <li><img src="images/purpledot.gif"> Incomplete</li>
                        <li><img src="images/reddot.gif"> Cancelled</li>
                    </ul>
                </div>
            </div>
            
        	<div class="pcCPsearchBox" id="cpBox6">
                <div class="CollapsiblePanel" id="cp6">
                  <div class="CollapsiblePanelTab"><h1>Payment Status Legend</h1></div>
                    <ul>
                        <li><img src="images/blueflag.gif"> Pending</li>
                        <li><img src="images/yellowflag.gif"> Authorized</li>
                        <li><img src="images/greenflag.gif"> Paid</li>
                        <li><img src="images/darkgreenflag.gif"> Refunded</li>
                        <li><img src="images/redflag.gif"> Voided</li>
                    </ul>
                </div>
            </div>

         <%
		 end if
		 
		IF lcase(section)<>"quickbooks" AND lcase(section)<>"ebay" AND lcase(pageTitle)<>"productcart ebay add-on" THEN
 
		 if session("admin")<>"0" and session("admin")<>"" then
		 %>
         

            
            <div class="pcCPsearchBox" id="cpBox2">
            	<form name="searchOrdersFooter" action="resultsAdvanced.asp?" class="pcForms">
                <div id="cp3" class="CollapsiblePanel">
                    <div class="CollapsiblePanelTab"><h1>Find an <strong>order</strong> by...</h1></div>
                    <div class="CollapsiblePanelContent">
                    <p><select class="select2" name="TypeSearch" size="1">
                        <option value="idOrder">Order ID</option>
                        <option value="orderCode">Order Code</option>
                        <% if GOOGLEACTIVE=-1 then %>
                        <option value="GoogleOrderID">Google Order ID</option>
                      <% end if %>
                        <option value="details">Product</option>
                        <option value="shipmentDetails">Shipping Type</option>
                        <option value="stateCode">State/Province Code</option>
                        <option value="CountryCode">Country Code</option>
                    </select></p>
                    <p><input class="textbox"  type="text" size="15" name="advquery" value="Enter Value" onFocus="clearText(this)"></p>
                    <p><input type="submit" name="B1" value="Find Orders" class="submit2">
                    <input type="button" value="More" onClick="location.href='invoicing.asp'"></p>
                    </div>
                 </div>
                 </form>
            </div>
        
        	<div class="pcCPsearchBox" id="cpBox1">
            	<form name="ajaxSearchFooter" method="post" action="srcPrds.asp?action=newsrc" class="pcForms">
                    <input type="hidden" name="referpage" value="NewSearch">
                    <input type="hidden" name="src_FormTitle1" value="Find Products">
                    <input type="hidden" name="src_FormTitle2" value="Product Search Results">
                    <input type="hidden" name="src_FormTips1" value="Use the following filters to look for products in your store.">
                    <input type="hidden" name="src_FormTips2" value="">
                    <input type="hidden" name="src_IncNormal" value="0">
                    <input type="hidden" name="src_IncBTO" value="0">
                    <input type="hidden" name="src_IncItem" value="0">
                    <input type="hidden" name="src_DisplayType" value="0">
                    <input type="hidden" name="pinactive" value="-1">
                    <input type="hidden" name="src_ShowLinks" value="1">
                    <input type="hidden" name="src_FromPage" value="LocateProducts.asp">
                    <input type="hidden" name="src_ToPage" value="">
                    <input type="hidden" name="src_Button2" value="Continue">
                    <input type="hidden" name="src_Button3" value="New Search">
                    <div id="cp1" class="CollapsiblePanel">
                        <div class="CollapsiblePanelTab"><h1>Find a <strong>product</strong> by...</h1></div>
                        <div class="CollapsiblePanelContent">
                        <p>SKU: <input name="sku" type="text" size="6" maxlength="150"></p>
                        <p>Keyword(s): <input type="text" name="keyWord" size="10"></p>
                        <p><select name="resultCnt" id="resultCnt">
                            <option value="5" selected>5</option>
                            <option value="10">10</option>
                            <option value="15">15</option>
                            <option value="20">20</option>
                            <option value="25">25</option>
                            <option value="50">50</option>
                            <option value="100">100</option>
                        </select>
                        <input type="hidden" name="act" value="newsrc">
                        <input name="Submit" type="submit" value="Go" class="submit2">
                        <input type="button" value="More" onClick="javascript:location.href='LocateProducts.asp';"></p>
                        </div>
                    </div>
			  </form>
            </div>
            
            <!--#include file="smallRecentProducts.asp"--> 
            
        	<div class="pcCPsearchBox" id="cpBox3">
            	<form name="listCustFooter" action="viewCustb.asp" class="pcForms">
                    <div id="cp4" class="CollapsiblePanel">
                        <div class="CollapsiblePanelTab"><h1>Find a <strong>customer</strong> by...</h1></div>
                        <div class="CollapsiblePanelContent">
                        <p>Last Name: <input type="text" name="key2" size="14" value=""></p>
                        <p>Company: <input type="text" name="key3" size="14" value=""></p>
                        <p>E-mail: <input type="text" name="key4" size="14" value=""></p>
                        <p><input type="submit" name="srcView" value="Search" class="submit2"></p>
                        <input type="hidden" name="key5" value="">
                        <input type="hidden" name="key6" value="">
                        </div>
                    </div>
				</form>
             </div>
             
        <% 
		 end if
		END IF
		%>
             
		</div>
	</div>
        
    <div id="pcFooter">
        <a href="about_terms.asp"><div style="float: left"><img src="images/pc_logo_100.gif" width="100" height="30" alt="ProductCart shopping cart software" border="0" /></div>Use of this software indicates acceptance of the End User License Agreement</a><br /><a href="http://www.productcart.com">Copyright&copy; 2001-<%=Year(now)%> NetSource Commerce. All Rights Reserved. ProductCart&reg; is a registered trademark of NetSource Commerce</a>.
    </div>

<script language="JavaScript" type="text/javascript">
	var MenuBar1 = new Spry.Widget.MenuBar("MenuBar1", {imgDown:"../includes/spry/images/SpryMenuBarDownHover.gif", imgRight:"../includes/spry/images/SpryMenuBarRightHover.gif"});

	var cookies = Spry.Utils.Cookie("read","multiplewidgets");
	
	if(cookies){
		cookies = cookies.split(",");
	}
	
	<%
	tmpStr=""
	IF lcase(section)<>"quickbooks" AND lcase(section)<>"ebay" AND lcase(pageTitle)<>"productcart ebay add-on" THEN
 
	if session("admin")<>"0" and session("admin")<>"" then
	tmpStr=tmpStr & "cp1.isOpen()"
	tmpStr=tmpStr & ",cp3.isOpen()"
	tmpStr=tmpStr & ",cp4.isOpen()"
	%>	
	var cp1 = new Spry.Widget.CollapsiblePanel("cp1", {contentIsOpen:cookies[0] && cookies[0] === 'false' ? false : true});
	var cp3 = new Spry.Widget.CollapsiblePanel("cp3", {contentIsOpen:cookies[2] && cookies[2] === 'false' ? false : true});
	var cp4 = new Spry.Widget.CollapsiblePanel("cp4", {contentIsOpen:cookies[3] && cookies[3] === 'false' ? false : true});
	<%end if
	END IF%>
	<% 
	if pcv_ShowSmallRecentProducts=1 then 
	if tmpStr<>"" then
		tmpStr=tmpStr & ","
	end if
	tmpStr=tmpStr & "cp2.isOpen()"%>
	var cp2 = new Spry.Widget.CollapsiblePanel("cp2", {contentIsOpen:cookies[1] && cookies[1] === 'false' ? false : true});
	<% 
	end if 
	%>
	<%
	if pcInt_ShowOrderLegend = 1 then
	if tmpStr<>"" then
		tmpStr=tmpStr & ","
	end if
	tmpStr=tmpStr & "cp5.isOpen()"
	tmpStr=tmpStr & ",cp6.isOpen()"
	%>
	var cp5 = new Spry.Widget.CollapsiblePanel("cp5", {contentIsOpen:cookies[4] && cookies[4] === 'false' ? false : true});
	var cp6 = new Spry.Widget.CollapsiblePanel("cp6", {contentIsOpen:cookies[5] && cookies[5] === 'false' ? false : true});
	<%
	end if
	%>
	
	Spry.Utils.addUnLoadListener(function(){
		// save the panels state to an Array
		<%if tmpStr<>"" then%>
		Spry.Utils.Cookie("create","multiplewidgets",[<%=tmpStr%>]);
		<%end if%>
	});
</script>

</body>
</html>