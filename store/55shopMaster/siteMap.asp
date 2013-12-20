<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Site Map" %>
<% section="" %>
<%PmAdmin=0%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/settings.asp"-->
<!--#include file="../includes/storeconstants.asp"--> 
<!--#include file="../includes/opendb.asp"-->
<!--#include file="../includes/adovbs.inc"-->
<!--#include file="../includes/stringfunctions.asp"-->
<!--#include file="../includes/languagesCP.asp" -->

<!--#include file="AdminHeader.asp"-->

	<div class="pcCPsiteMap">
	<table class="pcCPcontent" style="margin-left:20px;">
		<tr>
			<td valign="top" width="33%">
				<!-- FIRST COLUMN -->
				<ul>
					<li><a href="menu.asp">Start Page</a></li>
					<%if (not isNull(findUser(pcUserArr,1,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
					<li>Settings
						<ul>
							<li><a href="AdminSettings.asp">Store Settings</a></li>
							<li><a href="AdminSettings.asp?tab=3">Display Settings</a></li>
              				<li><a href="SearchOptions.asp">Search Settings</a></li>
							<li><a href="checkoutOptions.asp">Checkout Options</a></li>
							<li><a href="emailSettings.asp">E-mail Settings</a></li>
							<li><a href="AdminButtons.asp">Edit Store Buttons</a></li>
							<li><a href="AdminIcons.asp">Edit Store Icons</a></li>
							<li><a href="genCatNavigation.asp">Generate Navigation</a></li>
                            <li><a href="genLinksa.asp">Generate Store Links</a></li>
							<li><a href="genStoreMap.asp">Generate Store Map</a></li>
							<li><a href="blackout_main.asp">Manage Blackout Dates</a></li>
							<li><a href="GGG-GiftWrapOptions.asp">Manage Gift Wrapping</a></li>
							<li><a href="adminFBsettings.asp">Manage Help Desk</a></li>
							<li><a href="manageCountries.asp">Manage Countries</a></li>
							<li><a href="manageStates.asp">Manage States</a></li>
							<%if session("PmAdmin")="19" then%>
							<li><a href="AdminSecuritySettings.asp">Adv. Security Settings</a></li>
                            <li><a href="pcSecureKeyUpdate.asp">Update Encryption Key</a></li>
							<li><a href="AdminUserManager.asp">Manage Users</a></li>
							<li><a href="passwordchange.asp">Change User Login</a></li>
							<%end if%>
						</ul>
					</li>
					<%
					end if
					
					' Specials and discounts
					if (not isNull(findUser(pcUserArr,3,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
					<li>Marketing &amp; Discounts
						<ul>
							<li><a href="manageHomePage.asp">Manage Home Page</a></li>
							<li><a href="AdminFeatures.asp">Featured Products</a></li>
							<li><a href="manageBestSellers.asp">Manage Best Sellers</a></li>
							<li><a href="manageNewArrivals.asp">Manage New Arrivals</a></li>
                            <li><a href="manageRecentlyReviewed.asp">Manage Recently Reviewed</a></li>
							<li><a href="ManageSpecials.asp">Manage Specials</a></li>
							<li><a href="crossSellSettings.asp">Manage Cross Selling</a></li>
							<li><a href="AdminDiscounts.asp">Manage Electronic Coupons</a></li>
							<li><a href="ggg_managegcs.asp">Manage Gift Certificates</a></li>
                            <li><a href="PromotionPrdSrc.asp">Manage Promotions</a></li>
							<li><a href="viewDisca.asp">Quantity Discounts by Product</a></li>
							<li><a href="viewCatDisc.asp">Quantity Discounts by Category</a></li>
							<li><a href="RpStart.asp">Manage Reward Points</a></li>
							<li><a href="genSocialNetworkWidget.asp">Generate E-commerce Widget </a></li>
							<li><a href="genGoogleSiteMap.asp">Generate Google Sitemap</a></li>
						</ul>
					</li>
					<%end if

					' Shipping options
					if (not isNull(findUser(pcUserArr,4,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
					<li>Shipping
						<ul>
							<li><a href="modFromShipper.asp">Shipping Settings</a></li>
							<li><a href="viewShippingOptions.asp">Add or View Shipping Services</a></li>
							<li><a href="OrderShippingOptions.asp">Set Display Order</a></li>
							<li><a href="DeliveryZipCodes_main.asp">Set Delivery Zip Codes</a></li>
						</ul>
					</li>
					<%end if
					
					' Payment options
					if (not isNull(findUser(pcUserArr,5,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
					<li>Payments
						<ul>
							<li><a href="pcPaymentSelection.asp">Add New Option</a></li>
							<li><a href="PaymentOptions.asp">View/Modify Options</a></li>
							<li><a href="OrderPaymentOptions.asp">Set Display Order</a></li>
						</ul>
					</li>
					<%end if
					
					' Tax options
					if (not isNull(findUser(pcUserArr,6,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
					<li>Taxes
						<ul>
							<li><a href="AdminTaxSettings.asp">View/Edit Tax Options</a></li>
						</ul>
					</li>
					<%end if%>
					
					<li><a href="about_credits.asp">About ProductCart</a>
						<ul>
						<li><a href="about_credits.asp">Copyright &amp; Credits</a></li>
						<li><a href="about_terms.asp">Terms &amp; Conditions</a></li>
						<li><a href="help.asp">Technical Support</a></li>
						</ul>
					</li>					
				</ul>
				<!-- END FIRST COLUMN -->
				</td>
				<td valign="top" width="33%">
				<!-- SECOND COLUMN -->
				<ul>
					<li>Products
						<%if (not isNull(findUser(pcUserArr,2,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
						<ul>
							<li><a href="addProduct.asp?prdType=std">Add New Product</a></li>
							<li><a href="LocateProducts.asp?cptype=0">Locate a Product</a></li>
							<li><a href="manageOptions.asp">Manage Options</a></li>
							<li><a href="ManageCFields.asp">Manage Custom Fields</a></li>
							<li><a href="instCata.asp">Add Category</a></li>
							<li><a href="manageCategories.asp">Manage Categories</a></li>
							<li><a href="BrandsAdd.asp">Add Brand</a></li>
							<li><a href="BrandsManage.asp">Manage Brands</a></li>
							<li><a href="crossSellSettings.asp">Cross Selling</a></li>
							<li><a href="ImageUploada.asp">Upload Images</a></li>
							<li><a href="index_import_help.asp">Import Products</a></li>
							<li><a href="globalchanges.asp?nav=0">Global Changes</a></li>
							<li><a href="updPrdPrices.asp">Update Product Prices</a></li>
							<li><a href="viewStock.asp">Update Inventory</a></li>
							<li>Manage Suppliers
								<ul>
									<li><a href="sds_addnew.asp?pagetype=0">Add New Supplier</a></li>
									<li><a href="sds_manage.asp?pagetype=0">Find a Supplier</a></li>
									<li><a href="manageNewsWiz.asp?pagetype=0">Contact Suppliers</a></li>
								</ul>
							</li>
							<li>Manage Drop-Shippers
								<ul>
									<li><a href="sds_addnew.asp?pagetype=1">Add New Drop-Shipper</a></li>
									<li><a href="sds_manage.asp?pagetype=1">Find a Drop-Shipper</a></li>
									<li><a href="manageNewsWiz.asp?pagetype=1">Contact Drop-Shippers</a></li>
								</ul>
							</li>					
							<li>Product Reviews
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
							<%if session("PmAdmin")="19" then%>
							<li><a href="purgeproducts.asp">Purge deleted product</a></li>
							<li><a href="purgeallproducts.asp">Purge all products, categories, &amp; related orders</a></li>
							<%end if%>
						</ul>
						<%end if%>
					</li>
					
					<% if scBTO=1 then
							if (not isNull(findUser(pcUserArr,2,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
							<li>BTO Products
								<ul>
									<li><a href="BTOStart.asp">About BTO Products</a></li>
									<li><a href="BTOSettings.asp">BTO Settings</a></li>
									<li><a href="addProduct.asp?prdType=bto">Add New BTO Product</a></li>
									<li><a href="LocateProducts.asp?cptype=1">Locate a BTO Product</a></li>
									<li><a href="addProduct.asp?prdType=item">Add New BTO Item</a></li>
									<li><a href="LocateProducts.asp?cptype=2">Locate a BTO Item</a></li>
									<li><a href="updBTOPrdPrices.asp">Update Base Prices</a></li>
									<li><a href="updBTOiPrdPrices.asp">Update BTO Item Prices</a></li>
									<li><a href="updateBTOprices.asp">Update Configuration Prices</a></li>
									<li><a href="AddRmvBTOItemsMulti1.asp">Assing/Remove items to/from Multiple BTO Products</a></li>
									<li><a href="ApplyBTOCatMulti1.asp">Update Category Settings in Multiple BTO Products</a></li>
									<li><a href="globalchanges.asp?nav=1">Global Changes</a></li>
								</ul>
							</li>
						<% end if
					end if %>
					</ul>
					<!-- END SECOND COLUMN -->
					</td>
					
					<td valign="top" width="33%">
					<!-- THIRD COLUMN -->
						<ul>
						<%
					
						' Orders
						if (not isNull(findUser(pcUserArr,9,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
						<li>Orders
							<ul>
								<li><a href="invoicing.asp">Locate an Order</a></li>
								<li><a href="resultsAdvancedAll.asp?B1=View+All&dd=1">View All Orders</a></li>
								<li><a href="viewCusta.asp">View Orders by Customer</a></li>
								<li><a href="resultsAdvancedAll.asp?B1=View+All&dd=1&OType=1">View Incomplete Orders</a></li>
								<li><a href="batchprocessorders.asp">Batch Process Orders</a></li>
								<li><a href="batchshiporders.asp">Batch Ship Orders</a></li>
								<li><a href="creditCardPurge_index.asp">Purge Credit Card Numbers</a></li>
								<li><a href="adminviewallmsgs.asp">Help Desk: View All Postings</a></li>
							</ul>
						</li>
						<%end if
					
						' Reports and exports
						if (not isNull(findUser(pcUserArr,10,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
						<li>Reports
							<ul>
								<li><a href="srcOrdByDate.asp">View Sales Reports</a></li>
								<li><a href="exportData.asp">Custom Data Export</a></li>
								<li><a href="exportFroogle.asp">Export to Google Shopping</a></li>
								<% If scBTO=1 then %>
									<li><a href="srcQuotes.asp">View Quotes</a></li>
								<% end if %>
							</ul>
						</li>
						<%end if
						
							' Customers
							if (not isNull(findUser(pcUserArr,7,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
							<li>Customers
								<ul>
									<li><a href="viewCusta.asp">Locate a Customer</a></li>
               						<li><a href="viewCustb.asp?mode=LAST">View Latest Customers</a></li>
									<li><a href="instCusta.asp">Add New Customer</a></li>
									<li><a href="AdminCustomerCategory.asp">Manage Pricing Categories</a></li>
									<li><a href="manageCustFields.asp">Manage Special Fields</a></li>
									<li><a href="custindex_import_help.asp">Import Customers</a></li>
									<li><a href="manageNewsWiz.asp">Newsletter Wizard</a></li>
								</ul>
							</li>
							<%end if
							
							' Affiliates
							if (not isNull(findUser(pcUserArr,8,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
							<li>Affiliates
								<ul>
                					<li><a href="pcAffiliateSettings.asp">Affiliate Settings</a></li>
									<li><a href="instAffa.asp">Add New Affiliate</a></li>
									<li><a href="AdminAffiliates.asp">View/Modify Affiliates</a></li>
									<%if (instr(session("PmAdmin"),"10*")>0) or (session("admin")=0) or (session("PmAdmin")="19") then%>
										<li><a href="srcOrdByDate.asp#aff">View Affiliate Sales</a></li>
									<%end if%>
								</ul>
							</li>
							<%end if
                            
                           	' Pages
							if (not isNull(findUser(pcUserArr,11,pcUserArrCount))) or (not isNull(findUser(pcUserArr,12,pcUserArrCount))) or (session("admin")=0) or (session("PmAdmin")="19") then%>
							<li>Pages
								<ul>
                					<li><a href="cmsManage.asp">Manage Content Pages</a></li>
									<li><a href="cmsAddEdit.asp">Add New</a></li>
									<li><a href="cmsNavigation.asp">Add/Edit Navigation</a></li>
								</ul>
							</li>
							<%end if %>
                            <li><a href="about_credits.asp">About ProductCart</a>
                                <ul>
                                <li><a href="about_credits.asp">Copyright &amp; Credits</a></li>
                                <li><a href="about_terms.asp">Terms &amp; Conditions</a></li>
                                <li><a href="help.asp">Technical Support</a></li>
                                </ul>
                            </li>
                            <li><a href="logoff.asp">Log Off</a></li>
                            <li><a href="help.asp">Help</a></li>
						</ul>
						<!-- END THIRD COLUMN -->
			</td>
		</tr>
	</table>
	</div>

<!--#include file="AdminFooter.asp"-->