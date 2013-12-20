<ul id="MenuBar1" class="MenuBarHorizontal">
  <li class="firstLevel"><a class="MenuBarItemSubmenu" href="menu.asp">HOME</a></li>
<%
pcUserArr = split(session("PmAdmin"),"*")
pcUserArrCount = ubound(pcUserArr)-1

if (not isNull(findUser(pcUserArr,1,pcUserArrCount))) or (not isNull(findUser(pcUserArr,6,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
  <li class="firstLevel"><a class="MenuBarItemSubmenu" href="#">SETTINGS</a>
  	<ul>
<%
	if (not isNull(findUser(pcUserArr,1,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then %>
        <li><a class="MenuBarItemSubmenu" href="#">Store Settings</a>
            <ul>
                <li><a href="AdminSettings.asp">General Settings</a></li>
                <li><a href="AdminSettings.asp?tab=1">Company Details</a></li>
                <li><a href="AdminSettings.asp?tab=3">Display Settings</a></li>
                <li><a href="AdminSettings.asp?tab=2">Order &amp; Customer Settings</a></li>
				<li><a href="AdminSettings.asp?tab=5">Miscellaneous Settings</a></li>
				<%'Mobile-S%>
				<li><a href="MobileSettings.asp">Mobile Commerce Settings</a></li>
				<%'Mobile-E%>
                <li><a href="emailSettings.asp">E-mail Settings</a></li>
                <li><a href="SearchOptions.asp">Search Settings</a></li>
                <li><a href="checkoutOptions.asp">Checkout Options</a></li>
    			<li><a href="blackout_main.asp">Blackout Dates</a></li>
            </ul>
        </li>
        <li><a class="MenuBarItemSubmenu" href="#">Images &amp; Navigation</a>
            <ul>
                <li><a href="genCatNavigation.asp">Generate Category Navigation</a></li>
                <li><a href="genGoogleSiteMap.asp">Generate Sitemaps</a></li>
                <li><a href="genStoreMap.asp">Generate Store Map</a></li>
                <li><a href="genLinksa.asp">Get Store Links</a></li>
                <li><a href="pcv4_image_upload.asp">Upload Images</a></li>
                <li><a href="AdminButtons.asp">Upload Store Buttons</a></li>
                <li><a href="AdminIcons.asp">Upload Store Icons</a></li>
            </ul>
        </li>   
    	<li><a href="ggg-GiftWrapOptions.asp">Manage Gift Wrapping</a></li>
		<%if ((not isNull(findUser(pcUserArr,1,pcUserArrCount))) and (not isNull(findUser(pcUserArr,2,pcUserArrCount)))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
    	<li><a href="ggg_manageGCs.asp">Manage Gift Certificates</a></li>
        <%end if%>
        <%if ((not isNull(findUser(pcUserArr,1,pcUserArrCount))) and (not isNull(findUser(pcUserArr,7,pcUserArrCount))) and (not isNull(findUser(pcUserArr,9,pcUserArrCount)))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
    	<li><a href="ggg_manageGRs.asp">Manage Gift Registries</a></li>
        <%end if%>
    	<li><a href="adminFBsettings.asp">Manage Help Desk</a></li>
    	<li><a href="manageCountries.asp">Manage Countries</a></li>
    	<li><a href="manageStates.asp">Manage States</a></li>
	<%end if
	if (not isNull(findUser(pcUserArr,6,pcUserArrCount))) or (not isNull(findUser(pcUserArr,1,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
    	<li><a href="AdminTaxSettings.asp">Manage Taxes</a></li>
	<%
	end if
	if session("PmAdmin")="19" then%>
        <li><a href="AdminSecuritySettings.asp">Advanced Security Settings</a></li>
    	<li><a href="AdminUserManager.asp">Manage Users</a></li>
    	<li><a href="passwordchange.asp">Update Master User</a></li>
    	<li><a href="pcSecureKeyUpdate.asp">Update Encryption Key</a></li>        
    <%
	end if
	%>
  	</ul>
  </li>
<%
end if
if (not isNull(findUser(pcUserArr,2,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
  <li class="firstLevel"><a class="MenuBarItemSubmenu" href="#">PRODUCTS</a>
  	<ul>
	<%if (not isNull(findUser(pcUserArr,2,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
    	<li><a href="addProduct.asp?prdType=std">Add New Product</a></li>
    	<li><a href="LocateProducts.asp?cptype=0">Locate a Product</a></li>
    	<li><a href="index_import_help.asp">Import Products</a></li>
		<li><a href="iistep1.asp">Import Additional Product Images</a></li>
		<%
        'APP-S
        if instr(scSubVersion,"a")>0 then
        %>
            <li><a class="MenuBarItemSubmenu" href="#">Manage Apparel Products</a>
                <ul>
					<li><a href="app-settings.asp">Apparel Settings</a></li>
                    <li><a href="app-LocateProducts.asp">Locate an Apparel Product</a></li>
                    <li><a href="app-subPrdsBatch.asp">Batch Create Sub-Products</a></li>
                    <li><a href="index_import_help.asp">Import Apparel Products</a></li>
                    <li><a href="app-index_import_help.asp">Import Sub-Products</a></li>
                </ul>
            </li>
		<%
        end if
        'APP-E
        %>
        <li><a class="MenuBarItemSubmenu" href="#">Manage Categories</a>
            <ul>
                <li><a href="manageCategories.asp">View &amp; Edit Categories</a></li>
		    	<li><a href="instCata.asp">Add New Category</a></li>
		        <li><a href="catindex_import_help.asp">Import Categories</a></li>
		        <li><a href="ReverseCatImport_step1.asp">Export Categories</a></li>
                <li><a href="genCatNavigation.asp">Generate Category Navigation</a></li>
		        <li><a href="../pc/viewcategories.asp" target="_blank">Browse in the Storefront</a></li>
            </ul>
        </li>
        <li><a class="MenuBarItemSubmenu" href="#">Manage Brands</a>
            <ul>
                <li><a href="BrandsManage.asp">List Brands</a></li>
		    	<li><a href="BrandsAdd.asp">Add New Brand</a></li>
            </ul>
        </li>
        <li><a class="MenuBarItemSubmenu" href="#">Manage Custom Fields</a>
            <ul>
                <li><a href="ManageCFields.asp">Overview</a></li>
		    	<li><a href="ManageSearchFields.asp">Manage Custom Search Fields</a></li>
                <li><a href="addSFtoPrds.asp?nav=">Add Search Field to Products</a></li>
                <li><a href="addSFtoCats.asp">Add Search Field to Categories</a></li>
                <li><a href="addCFtoPrds.asp?nav=">Add Input Field to Products</a></li>
            </ul>
        </li>
        <li><a href="ggg_manageGCs.asp">Manage Gift Certificates</a></li>
        <li><a href="manageOptions.asp">Manage Product Options</a></li>
        <li><a class="MenuBarItemSubmenu" href="#">Manage Product Reviews</a>
            <ul>
                <li><a href="PrvSettings.asp">Product Reviews Settings</a></li>
                <li><a href="prv_ManageBadWords.asp">Bad Words Filter</a></li>
                <li><a href="prv_FieldManager.asp">Add/Edit Fields</a></li>
                <li><a href="prv_PrdExc.asp">Product Exclusions</a></li>
                <li><a href="prv_SpecialPrd.asp">Product-specific Settings</a></li>
                <li><a href="prv_ManageRevPrds.asp?nav=1">Pending Reviews</a></li>
                <li><a href="prv_ManageRevPrds.asp?nav=2">Live Reviews</a></li>
            </ul>
        </li>
		<% 'SB S %>
        <li><a class="MenuBarItemSubmenu" href="sb_Default.asp">Manage Subscriptions</a>
        	<ul>           
                <li><a href="sb_manageAcc.asp">SubscriptionBridge Activation</a></li>
                <li><a href="sb_Settings.asp">SubscriptionBridge Settings</a></li>   
            	<li><a href="sb_CreatePackages.asp">Add SB Package Link</a></li>
                <li><a href="sb_ViewPackages.asp">Modify SB Package Links</a></li>
                <li><a href="http://wiki.subscriptionbridge.com/cartintegration/productcart/" target="_blank">Help</a></li>
			</ul>
        </li>
		<% 'SB E %>
        <li><a class="MenuBarItemSubmenu" href="#">Manage Suppliers</a>
            <ul>
                <li><a href="sds_addnew.asp?pagetype=0">Add New Supplier</a></li>
                <li><a href="sds_manage.asp?pagetype=0">Locate a Supplier</a></li>
                <li><a href="manageNewsWiz.asp?pagetype=0">Contact Suppliers</a></li>
            </ul>
        </li>
        <li><a class="MenuBarItemSubmenu" href="#">Manage Drop-Shippers</a>
            <ul>
                <li><a href="sds_addnew.asp?pagetype=1">Add New Drop-Shipper</a></li>
                <li><a href="sds_manage.asp?pagetype=1">Locate a Drop-Shipper</a></li>
                <li><a href="manageNewsWiz.asp?pagetype=1">Contact Drop-Shippers</a></li>
            </ul>
        </li>
        <li><a class="MenuBarItemSubmenu" href="#">Update Multiple Products</a>
            <ul>
                <li><a href="globalchanges.asp?nav=0">Global Changes</a></li>
                <li><a href="viewStock.asp">Update Inventory Levels</a></li>
                <li><a href="updPrdPrices.asp">Update Product Prices</a></li>
            </ul>
        </li>
    <%
	end if
	%>
  	</ul>
  </li>
<%
end if
'BTO ADDON-S
If scBTO=1 then 
	if (not isNull(findUser(pcUserArr,2,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
		<li class="firstLevel"><a class="MenuBarItemSubmenu" href="#">BTO</a>
  			<ul>
                <li><a href="BTOStart.asp">About BTO Products</a></li>
                <li><a href="BTOSettings.asp">BTO Settings</a></li>
                <li><a href="addProduct.asp?prdType=bto">Add New BTO Product</a></li>
                <li><a href="LocateProducts.asp?cptype=1">Locate a BTO Product</a></li>
                <li><a href="addProduct.asp?prdType=item">Add New BTO Item</a></li>
                <li><a href="LocateProducts.asp?cptype=2">Locate a BTO Item</a></li>
                <li><a href="manageCategories.asp?nav=1">Manage Categories</a></li>
				<li><a class="MenuBarItemSubmenu" href="#">Update Multiple BTO Products</a>
                	<ul>
						<li><a href="AddRmvBTOItemsMulti1.asp">Assign/Remove Products &amp; Items</a></li>
						<li><a href="ApplyBTOCatMulti1.asp">Update Category Settings</a></li>
                		<li><a href="globalchanges.asp?nav=1">Global Changes</a></li>
					</ul>
				</li>
				<li><a class="MenuBarItemSubmenu" href="#">Update BTO Prices</a>
                	<ul>
                        <li><a href="updBTOPrdPrices.asp">Base Prices</a></li>
                        <li><a href="updBTODefaultPrices.asp">Default Prices</a></li>
                        <li><a href="updBTOiPrdPrices.asp">Item Prices</a></li>
                        <li><a href="updateBTOprices.asp">Configuration Prices</a></li>
					</ul>
				</li>
                <li><a href="updateBTOItemQty.asp">Update Item Inventory Levels</a></li>
			</ul>
		</li>
	<% end if
	end if
'BTO ADDON-E

'CMS
if (not isNull(findUser(pcUserArr,11,pcUserArrCount))) or (not isNull(findUser(pcUserArr,12,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then %>
	<li class="firstLevel"><a class="MenuBarItemSubmenu" href="#">PAGES</a>
		<ul>
        	<li><a href="cmsManage.asp">Manage Content Pages</a></li>
            <li><a href="cmsAddEdit.asp">Add New Content Page</a></li>
            <% if (not isNull(findUser(pcUserArr,11,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then %>
            <li><a href="cmsNavigation.asp">Generate Page Navigation</a></li>
            <% end if %>
		    <li><a href="../pc/viewPages.asp" target="_blank">Browse in the Storefront</a></li>
        </ul>
	</li>
<%
end if
'CMS-END

if (not isNull(findUser(pcUserArr,3,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then %>
	<li class="firstLevel"><a class="MenuBarItemSubmenu" href="#">MARKETING</a>
		<ul>
			<li><a class="MenuBarItemSubmenu" href="#">Manage Sales</a>
                <ul>
                    <li><a href="sm_addedit_S1.asp?a=new">Create a New Sale</a></li>
                    <li><a href="sm_manage.asp">View &amp; Edit Pending Sales</a></li>
                    <li><a href="sm_start.asp">Start a Sale</a></li>
                    <li><a href="sm_stop.asp">Stop a Sale</a></li>
                    <li><a href="sm_sales.asp">Current &amp; Completed Sales</a></li>
                </ul>
            </li>
            <li><a class="MenuBarItemSubmenu" href="#">Manage Special Pages</a>
                <ul>
                    <li><a href="manageHomePage.asp">Home Page</a></li>
                    <li><a href="AdminFeatures.asp">Featured Products</a></li>
                    <li><a href="manageBestSellers.asp">Best Sellers</a></li>
                    <li><a href="manageNewArrivals.asp">New Arrivals</a></li>
                    <li><a href="manageRecentlyReviewed.asp">Recently Reviewed</a></li>
                    <li><a href="manageSpecials.asp">Specials</a></li>
                </ul>
            </li>
            <li><a class="MenuBarItemSubmenu" href="#">Manage Cross Selling</a>
                <ul>
                    <li><a href="crossSellSettings.asp?idmain=1">Cross Selling Settings</a></li>
                    <li><a href="crossSellView.asp">Existing Relationships</a></li>
                    <li><a href="crossSellAdd.asp">Add New Relationship</a></li>
                </ul>
            </li>
        	<li><a href="AdminDiscounts.asp">Manage Discount Codes (Coupons)</a></li>
        	<li><a href="ggg_managegcs.asp">Manage Gift Certificates</a></li>
			<% if (not isNull(findUser(pcUserArr,7,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then %>
        	<li><a href="ggg_manageGRs.asp">Manage Gift Registries</a></li>
            <% end if %>
        	<li><a href="PromotionPrdSrc.asp">Manage Promotions</a></li>
        	<li><a href="RpStart.asp">Manage Reward Points</a></li>
            <li><a class="MenuBarItemSubmenu" href="#">Manage Tiered Pricing</a>
                <ul>
        			<li><a href="viewDisca.asp">Quantity Discounts by Product</a></li>
        			<li><a href="viewCatDisc.asp">Quantity Discounts by Category</a></li>
                </ul>
            </li>
            <li><a class="MenuBarItemSubmenu" href="#">Generate File for...</a>
                <ul>
                    <li><a href="exportBing.asp">Bing Shopping</a></li>
                    <li><a href="exportFroogle.asp">Google Shopping</a></li>
                    <li><a href="pcNextTag_step1.asp">NexTag </a></li>
                    <li><a href="pcYahoo_step1.asp">Yahoo!</a></li>
                    <li><a href="genSocialNetworkWidget.asp">E-Commerce Widget</a></li>
                </ul>
            </li>
        	<li><a href="eb_home.asp">Publish to eBay</a></li>
		</ul>
	</li>
<%
end if
if (not isNull(findUser(pcUserArr,4,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then
%>
	<li class="firstLevel"><a class="MenuBarItemSubmenu" href="#">SHIPPING</a>
		<ul>
        	<li><a href="modFromShipper.asp">Shipping Settings</a></li>
        	<li><a href="viewShippingOptions.asp">Add or View Shipping Services</a></li>
        	<li><a href="OrderShippingOptions.asp">Set Display Order</a></li>
        	<li><a href="DeliveryZipCodes_main.asp">Set Delivery Zip Codes</a></li>
		</ul>
	</li>
<%
end if
if (not isNull(findUser(pcUserArr,5,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then
%>
	<li class="firstLevel"><a class="MenuBarItemSubmenu" href="#">PAYMENTS</a>
		<ul>
        	<li><a href="pcPaymentSelection.asp">Add New Payment Option</a></li>
        	<li><a href="PaymentOptions.asp">View/Modify Payment Options</a></li>
        	<li><a href="OrderPaymentOptions.asp">Set Display Order</a></li>
            <li><a href="AdminPaymentMerchantAccount.asp">Merchant Accounts</a></li>
            <li><a href="http://www.earlyimpact.com/gateway/" target="_blank">NetSource Commerce payment gateway</a></li>
		</ul>
	</li>
<%
end if
if (not isNull(findUser(pcUserArr,7,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then
%>
	<li class="firstLevel"><a class="MenuBarItemSubmenu" href="#">CUSTOMERS</a>
		<ul>
        	<li><a href="viewCusta.asp">Locate a Customer</a></li>
        	<li><a href="instCusta.asp">Add New Customer</a></li>
            <% if (session("admin")=0) or (session("PmAdmin")="19") then %>
        	<li><a href="AdminCustomerCategory.asp">Manage Pricing Categories</a></li>
            <% end if %>
        	<li><a href="manageCustFields.asp">Manage Special Fields</a></li>
        	<li><a href="viewCusta.asp">Place Order (Existing Customer)</a></li>
        	<li><a href="placeOrder.asp">Place Order (New Customer)</a></li>
        	<li><a href="custindex_import_help.asp">Import Customers</a></li>
        	<li><a href="manageNewsWiz.asp">Newsletter Wizard</a></li>
        	<li><a href="mailup_home.asp">MailUp Integration</a></li>
		</ul>
	</li>
<%
end if
if (not isNull(findUser(pcUserArr,8,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then
%>
	<li class="firstLevel"><a class="MenuBarItemSubmenu" href="#">AFFILIATES</a>
		<ul>
        	<% if session("PmAdmin")="19" then %>
        	<li><a href="pcAffiliateSettings.asp">Affiliate Settings</a></li>
            <% end if %>
        	<li><a href="instAffa.asp">Add New Affiliate</a></li>
        	<li><a href="AdminAffiliates.asp">View/Modify Affiliates</a></li>
        	<li><a href="srcOrdByDate.asp#aff">View Affiliate Sales</a></li>
		</ul>
	</li>
<%
end if
if (not isNull(findUser(pcUserArr,9,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then
%>
	<li class="firstLevel"><a class="MenuBarItemSubmenu" href="#">ORDERS</a>
		<ul>
        	<li><a href="invoicing.asp">Locate an Order</a></li>
        	<li><a href="resultsAdvancedAll.asp?B1=View+All&dd=1">View All Orders</a></li>
        	<li><a href="viewCusta.asp">View Orders by Customer</a></li>
        	<li><a href="resultsAdvancedAll.asp?B1=View+All&dd=1&OType=1">View Incomplete Orders</a></li>
            <li><a href="resultsAdvancedAll.asp?B1=View+All&dd=1&pcIntArchived=1">View Archived Orders</a></li>
        	<li><a href="batchprocessorders.asp">Batch Process Orders</a></li>
        	<li><a href="batchshiporders.asp">Batch Ship Orders</a></li>
        	<li><a href="creditCardPurge_index.asp">Purge Credit Card Numbers</a></li>
            <% IF trim(scGoogleAnalytics)<>"" AND NOT IsNull(scGoogleAnalytics) THEN %>
            <li><a href="pcGA_refund.asp">Google Analytis Adjustments</a></li>
            <% END IF %>
        	<li><a href="adminviewallmsgs.asp">Help Desk: View All Postings</a></li>
		</ul>
	</li>
<%
end if
if (not isNull(findUser(pcUserArr,10,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then
%>
	<li class="firstLevel"><a class="MenuBarItemSubmenu" href="#">REPORTS</a>
		<ul>
        	<li><a href="srcOrdByDate.asp">View Sales Reports</a></li>
			<li><a href="dashboard.asp">Sales Charts and Graphs</a></li>
        	<li><a href="CheckoutReport.asp">View Drop-Off/Conversions</a></li>
			<li><a href="viewSCLogs.asp">View Saved Carts Statistics</a></li>
        	<li><a href="srcOtherReports.asp">View Other Reports</a></li>
			<% 'BTO ADDON-S
            If scBTO=1 then %> 
        	<li><a href="srcQuotes.asp">View &amp; Edit Quotes</a></li>
			<% end if
            'BTO ADDON-E %>
        	<li><a href="exportData.asp">Custom Data Exports</a></li>
        	<li><a href="qb_home.asp">Synchronize with QuickBooks</a></li>
			<%if session("PmAdmin")="19" then%>
        	<li><a href="viewCPLogs.asp">View Control Panel Logs</a></li>
        	<li><a href="XMLToolsManager.asp">XML Tools Manager</a></li>
            <%end if%>
		</ul>
	</li>
<%
end if
%>
	<li class="firstLevel"><a class="MenuBarItemSubmenu" href="#">HELP</a>
		<ul>
        	<li><a href="help.asp">Help Center</a></li>
            <li><a href="https://www.earlyimpact.com/eistore/productcart/pc/custpref.asp" target="_blank">Submit a Ticket</a></li>
        	<li><a href="helpTwitter.asp">Twitter Updates</a></li>
        	<li><a href="helpErrorFinder.asp">Error Information</a></li>
        	<li><a href="pcTSUtility.asp">Troubleshooting Utility</a></li>
			<%if session("PmAdmin")="19" then%>
			<li><a class="MenuBarItemSubmenu" href="#">Database Clean Up Tool</a>
			<ul>
			   	<li><a href="PurgeCustSessions.asp">Remove customer sessions</a></li>
				<li><a href="PurgeSavedCarts.asp">Remove saved carts</a></li>
			</ul>
            <li><a href="http://www.earlyimpact.com/productcart/addons.asp" target="_blank">Extend ProductCart</a></li>
        	<li><a href="http://www.earlyimpact.com/productcart/support/updates/pcvm_listPatches.asp?keyid=<%=scCrypPass%>&vnum=<%=scVersion%>&svnum=<%=scSubVersion%>&sp=<%=scSP%>" target="_blank">Check for Updates &gt;&gt;</a></li>
			<%end if%>
		</ul>
	</li>
	<li class="firstLevel"><a class="MenuBarItemSubmenu" href="logoff.asp">EXIT</a></li>
</ul>