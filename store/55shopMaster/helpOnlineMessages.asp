<style>
	ul {
		padding-top: 5px;
	}
	li {
		padding-bottom: 5px;
	}
</style>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2013. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'--------------------------------------------
' ONLINE HELP
' Help messages used in the Control Panel
'--------------------------------------------

' 100 = Drop shiping and Supplier Management
' 200 = Marketing features
' 300 = Payment options and order processing
' 400 = General administration tools
' 500 = Build To Order

' Example of syntax to launch Online Help:
' &nbsp;<a href="JavaScript:win('helpOnline.asp?ref=400')"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature"></a>


	select case pcIntHelpTip
	
	' Drop Shipping and Supplier Management
	
		case 101
			pcStrTitle = "Setting a Supplier as a Drop Shipper"
			pcStrDetails = "If this supplier ships products for you, check this option and the system will recognize it as a Drop Shipper. You will see the company listed whenever Drop Shippers are listed. Refer to the User Guide for more information on drop shipping."
			
		case 102
			pcStrTitle = "Ship-From Address"
			pcStrDetails = "The address from which products will be shipped by the drop-shipping company, which might be different from the billing address."
			
		case 103
			pcStrTitle = "Drop-Shipper Login Information"
			pcStrDetails = "The User Name and Password that the Drop Shipper can use to log into a special section of your store where order information can be updated. Once logged in, the Drop Shipper can update you on the status of the drop shipping requests. For example, the Drop Shipper could let you know that all products that were part of order XXXX were shipped, or that one or more products could not be shipped and indicate why."
			
		case 104
			pcStrTitle = "Drop Shipping settings: Notify customer when order is updated"
			pcStrDetails = "When this feature is active, if the Drop Shipper logs in and updates an order (e.g. all products have been shipped), the store manager is notified, <strong>AND</strong> the Order Shipped email is <strong>automatically</strong> sent to the customer. Otherwise, the store manager is notified, but the customer is <strong>not</strong> contacted."
			
		case 105
			pcStrTitle = "Drop Shipping settings: Only notify manually"
			pcStrDetails = "Any drop-shipper involved in an order is notified automatically via e-mail when the order status is updated from &quot;Pending&quot; to &quot;Processed&quot;, unless this option has been checked. If this option has been checked, then the drop-shipper is notified only when the store administrator uses the &quot;Send Drop-Shipper Notification&quot; feature, which allows the store administrator to send or re-send the notification e-mail at any time, independently of the processing of the order.<br /><br />Note that if the processing of the order happens automatically (e.g. order was placed using a payment option that had been setup for automatic order processing), then the notification e-mail is automatically sent to the drop-shipper unless the &quot;Only Notify Manually&quot; setting is turned on."
			
		case 106
			pcStrTitle = "Drop Shipping settings: Order notification e-mail"
			pcStrDetails = "The e-mail address to which order information is sent when a Drop Shipper is notified of products that need to be drop-shipped."
			
		case 107
			pcStrTitle = "Drop Shipping settings: Order notification content"
			pcStrDetails = "You can include either &quot;Products + Customer shipping information&quot; (shipping name and address) or &quot;Products Only&quot;. If you select the second option, no customer information is included. This assumes that you wish products to be shipped to you. Therefore, your store's shipping address is included instead of the customer's address)"
			
		case 108
			pcStrTitle = "Store Settings: Allowing separate shipments"
			pcStrDetails = "If you allow customers to purchase products that are back-ordered, there might be cases in which they order a combination of products that can be shipped immediately and products that cannot be shipped right away because they are back-ordered. In this scenario, you can allow your customers to tell you whether they would like to receive one shipment (they would wait until all products are available) or multiple shipments (products that can be shipped are shipped immediately).<br /><br />Note that it is up to you to decide whether additional shipping charges should occur in case the customer opts to receive the packages in multiple shipments. You can use the Edit Order feature to increase the shipping charges, if you wish to do so. You should provide your customers with details on your shipping policy in a &quot;Customer Service&quot; area of your store.<br /><br />You can notify customers that they are purchasing products that might be shipped in separate shipments by setting that option on the &quot;<a href='modFromShipper.asp' target='_blank'>Shipping Settings</a>&quot; page."
			
	' Marketing features
	
		case 201
			pcStrTitle = "Products to show: about &quot;out of stock&quot; products"
			pcStrDetails = "Note that when the &quot;Disregard Stock&quot; option has been selected for a product, the product is shown even if this setting is set to &quot;No&quot; and the product is out of stock."
			
		case 202
			pcStrTitle = "Products to show: about &quot;not for sale&quot; products"
			pcStrDetails = "You can set a products to be &quot;Not for Sale&quot; when adding or editing it. You can also set this property on multiple products at once using the <a href='globalchanges.asp?nav=0' target='_blank' onClick='JavaScript: window.close()'>Manage Products &gt; Global Changes</a> features"
			
		case 203
			pcStrTitle = "Home page: highlighting the first featured product"
			pcStrDetails = "You can make the first featured product more visible on your store's home page by using this feature (e.g. 'Product of the Month'). The order in which products are shown is set on the '<a href='AdminFeatures.asp' target='_blank' onClick='JavaScript: window.close()'>Manage Featured Products</a>' page. You can indicate which product is the 'first' featured product by using the 'Order' field on the page."
			
		case 204
			pcStrTitle = "Search Engine Optimization: custom Meta Tags"
			pcStrDetails = "You can enter product- and category-specific Meta Tags to provide search engines with more accurate information on what is found on the page. Although they are no longer the main element used by search engines to rank Web sites, well written Title and Description meta tags can certainly contribute to good search engine rankings.<br /><br />The Meta Tags that you enter on this page will be written to the corresponding category or product page in your storefront <u>only</u> if the file &quot;<strong>include-metatags.asp</strong>&quot; is used in &quot;<strong>pc/header.asp</strong>&quot;. Please refer to the ProductCart User Guide for information about customizing the graphical interface of your store (header.asp and footer.asp).<br /><br />If you have a large number of products and/or categories and cannot write Meta Tags for each for them, don't worry. Just make sure that you use the file &quot;include-metatags.asp&quot; in your version of &quot;pc/header.asp&quot;, and ProductCart will dynamically create the Meta Tags for you by using the product name for the Title tag and a portion of the product description for the Description tag.<br /><br />If you are interested in learning more about Meta Tags, you will find a lot of articles on the Internet. For example, see the <a href='http://blog.earlyimpact.com/2011/01/google-seo-starter-guide-and.html' target='_blank'>Google SEO Starter Guide &gt;&gt;</a>."
			
		case 205
			pcStrTitle = "Cross Selling: Bundle vs. Not a Bundle"
			pcStrDetails = "You can create two types of cross-selling relationships for your products:<ul><li><strong>Bundle</strong>. When you set up the cross-selling relationship as a bundle, ProductCart will display the two products (main product + cross-sold item) in a special area of the page, as a bundle. Customers can <u>either</u> add to the cart the main product <u>or</u> the bundled products.<br /><br />Since you can apply a discount to the bundle, this feature allows you to promote the sale of bundled products, thus helping you increase the Average Order Amount, which is a great way to increase sales on your store.<br /><br /></li><li><strong>Not a Bundle</strong>. Choose this option if you simply want to promote cross-selling items, without proposing that they are purchased with the main product as a bundle (e.g. a bundle might not make sense to the customer). These products are shown separately from bundled products on the product details page. Multiple cross-sold products can be added to the cart together with the main product.</li></ul>"
			
		case 206
			pcStrTitle = "Cross Selling: Bundle Discount"
			pcStrDetails = "When you set the type of the cross-selling relationship to &quot;Bundle&quot;, you can promote the purchase of the bundled products by offering your customers a discount. The discount can be a set amount or a percentage off the sum of the two product prices. Check the % checkbox if you want the value to be applied as a percentage of the price."
			
		case 207
			pcStrTitle = "Cross Selling: Required Accessory"
			pcStrDetails = "When you set the type of the cross-selling relationship to &quot;Not a Bundle&quot;, you can specify whether the cross-sold item is required. This allows you to force the customer to purchase an accessory that is required based on the main product that the customer is purchasing."
			
		case 208
			pcStrTitle = "Affiliate Settings: Save Affiliate ID"
			pcStrDetails = "When this option is active, a harmless cookie with the affiliate's unique ID is saved on the customer's computer the first time that customer visits your store after clicking on link to the store that included the affiliate ID.<br /><br />If the same customer later returns to place an order (or place another order), the system will recognize the original referring affiliate and store the corresponding affiliate ID with the new order.<br /><br />This allows you to let affiliates earn a commission on a sale even if the referred customer does not purchase during the first visit to your store.<br /><br />You can select the number of days the cookie will remain active on the customer's computer before it expires. For example, enter 365 if you want the cookie to remain valid for 1 year."

		case 209
			pcStrTitle = "Affiliate Settings: Maximum Number of Orders Per Customer"
			pcStrDetails = "By setting the maximum number of orders, you can limit how many times an affiliate can collect commisions for referring a certain customer. Note that the &quot;Save Affiliate ID&quot; option must be active for commissions to be earned by an affiliate beyond the first order placed by the referred customer (otherwise the system will not know who the referring affiliate was when the customer places a new order)."

		case 210
			pcStrTitle = "Affiliate Settings: Automatic vs. Manual Approval"
			pcStrDetails = "By default new affiliates are not automatically approved. Their affiliate account remains pending until you review it and approve it from the Control Panel. If you prefer, you can enable automatic approval by using this option.<br /><br />In either case you will receive a notification e-mail when a new affiliate signs up.<br /><br />Affiliates will not be able to log into their account until they have been approved (e.g. to obtain store links that include their affiliate ID)."	
			
		case 211
			pcStrTitle = "Affiliate Settings: Default Commission"
			pcStrDetails = "Enter the commission that you normally give your affiliates: it will be saved automatically to their affiliate account when a new account is created. You can edit any affiliate at any time to change the commission value, if needed. For a 20% commission, enter 20.<br /><br />If no default commission is speficied, it will default to zero. However, if you have your affiliate program setup to auto-approve affiliates, this will mean that affiliate accounts will be created with a 0% commission."		
			
		case 212
			pcStrTitle = "Affiliate Settings: Exclude Wholesale Customers"
			pcStrDetails = "When this feature is turned on, an affiliate will not earn any commissions on an order placed by a wholesale customer. Even if the customer clicked on an affiliate link before purchasing the product, the Affiliate ID associated with that link will not be saved to the database if the customer is a wholesale customer or if the customer belongs to a Pricing Category with wholesale privileges."			
			
			
		case 213
			pcStrTitle = "Microsoft Bing Cashback: Product Title"
			pcStrDetails = "The <strong>Title</strong> should be an accurate representation of your product. For optimal effectiveness, the Title you submit in your data feed may need to be revised from its original appearance on your site. Examples of product types: <br /><br />For products such as Appliances, Computers, Electronics, etc.:<br /><br /><strong>Manufacturer/Brand >>> Model Number >>> Product Category</strong><br />Gateway >>> GM5626 >>> Desktop Computer<br />Nikon >>> Coolpix L15 >>> Digital Camera<br /><br />For media products such as Books, Music, & Movies:<br /><br /><strong>Title >>> Author/Artist/Director >>> Format</strong><br />Touch of Evil >>> Kay Hooper >>> Paperback<br />Saving Private Ryan >>> Steven Spielberg >>> Blu-Ray DVD<br /><br />We recommend that you use all three drop-down fields to create the product title, but only the first drop-down is required.<br /><br />Note: ProductCart will remove any special characters not allowed by the Cashback system. Also, any Title that excedes the maximum number of characters will be truncated automatically. If you select an option and it is empty (e.g. you select a search field with no value for a certain product) ProductCart will substitute the Product Name for the Title."	
			
		case 214
			pcStrTitle = "Promotion Settings: Promotion Message"
			pcStrDetails = "The &quot;Promotion Message&quot; is the message that will be displayed when the product being promoted is shown in the storefront. The message is shown in the area of the product details page where prices are shown. The message is also shown in a special region of the 'View Shopping Cart' page. <br><br>Note that the message does not automatically include the product name (you have complete flexibility in what the message says), so in most cases you will want to include it in the message itself. Here are a couple of examples.<br><br><em>Buy 2 &quot;Deluxe Widgets&quot; and get the 3rd one free.</em><br><em>Buy 1 pair of &quot;SuperRunner Pro&quot; and get the next pair half off.</em><br><br>For more information on this feature, see the <a href='http://wiki.earlyimpact.com/productcart/marketing-promotions' target='_blank'>ProductCart documentation</a>."
			
		case 215
			pcStrTitle = "Promotion Settings: Confirmation Message"
			pcStrDetails = "The &quot;Confirmation Message&quot; is the message that is shown on the shopping cart page when a promotion is applied to the shopping cart contents.<br><br>Note that the message does not automatically include the product name (you have complete flexibility in what the message says), so in most cases you will want to include it in the message itself. Here are a couple of examples.<br><br><em>Buy 2 get 1 free promotion on &quot;Deluxe Widgets&quot;.</em><br><em>Total reflects 50% off 2nd pair of &quot;SuperRunner Pro&quot;.</em><br><br>For more information on this feature, see the <a href='http://wiki.earlyimpact.com/productcart/marketing-promotions' target='_blank'>ProductCart documentation</a>."	
			
		case 216
			pcStrTitle = "Promotion Settings: Short Description"
			pcStrDetails = "The &quot;Short Description&quot; is the message that is saved to the database together with the other order details and shown on all order details pages together with other discounts and promotions that apply to the order. The customer knows what was purchased and which promotion was running at the time, so the message can be shorter than the ones shown in the storefront, prior to the purchase.<br /><br />It should describe the promotion that has been applied to the order in a way that can be easily recognized at a later time. <br><br>Note that the message does not automatically include the product name (you have complete flexibility in what the message says), so in most cases you will want to include it in the message itself. Here are a couple of examples.<br><br><em>Buy 2 get 1 free on &quot;Deluxe Widgets&quot;</em><br><em>50% off 2nd pair of &quot;SuperRunner Pro&quot;</em>.<br><br>For more information on this feature, see the <a href='http://wiki.earlyimpact.com/productcart/marketing-promotions' target='_blank'>ProductCart documentation</a>."
			
		case 217
			pcStrTitle = "Promotion Settings: Apply to Next N Units"
			pcStrDetails = "This is the setting that allows you to implement promotions such as <strong>Buy One, Get One Free</strong>: in this case (and for most promotions you will create) you will set the promotion so that it applies to the next unit. In a &quot;Buy One, Get One Free&quot; promotion you will set <strong>Apply to next N Units</strong> to 1 and the discount percentage to 100%.<br><br>However, N might be higher than 1. For example, you could create a promotion that says: &quot;Buy 2 bottles of &quot;Nice Wine&quot; and get 30% off the next 2 bottles&quot; (here <strong>N=2</strong>). Or you may want to liquidate some inventory with a promotion such as &quot;Buy 5 and get 5 free&quot; (here <strong>N=5</strong>).<br><br>If <strong>N=0</strong> the promotion will be applied to all additional units. For example: &quot;Buy 2 pairs of 'Red &amp; Blue Jeans' and get 20% off any additional pair&quot;.<br><br>For more information on this feature, see the <a href='http://wiki.earlyimpact.com/productcart/marketing-promotions' target='_blank'>ProductCart documentation</a>."	
			
		case 218
			pcStrTitle = "Promotion Settings: Quantity Trigger"
			pcStrDetails = "The promotion might require more than one unit being purchased in order for it to apply to the order. This is where you can specify that number. For example:<br><br> <strong>Buy 2 Get 1 Free</strong> will require the 'Quantity Trigger' to be set to 2.<br><strong>Buy 1 pair get 50% the 2nd pair</strong> will require the 'Quantity Trigger' to be set to 1."	
			
		case 219
			pcStrTitle = "Promotion Settings: Applicability"
			pcStrDetails = "The promotion by default is set to apply to more units of the same product (product-based promotion) or more units of the products contained in the selected category (category-based promotion). So when you first create a promotion you will find the selected product or the selected category respectively in the Products or Categories filter.<br /><br />For example, a promotion such as &quot;<em>Buy one and get 50% off the second one</em>&quot; is meant to push higher sales of the same item (e.g. buy one silk tie and get the next one 50% off).<br /><br />However, the promotion could certainly apply to other products or categories of products. The product and category applicability parameters allow you to specify which products and/or categories of products the promotion applies to.<br /><br />For example, consider the following two different promotions:<ul><li>Buy one suit and get 50% off a certain group of dress shirts (which you specify in the <em>Products</em> filter).</li><li>Buy one suit and get 50% off all dress shirts (which you specify by adding the 'Dress Shirts' category to the <em>Categories</em> filter).</li></ul>NOTE: if no product and/or category applicability filters are set (i.e. you remove even the product or category filter that is initially set by default when you create a new promotion), then the promotion applies to everything in the store (e.g. buy 2 units of this product and get 10% off anything else in the store)."	
			
			
	' Payment Options and Order Processing
			
		case 301
			pcStrTitle = "Process orders when they are placed"
			pcStrDetails = "In ProductCart v3 and above you have full control on whether or not an order is processed automatically when it is placed, <u>regardless of the payment option used and how it has been setup</u>.<br /><br />If you want to have the order processed at the time it is placed, check this checkbox.<br /><br />When an order is processed, a number of tasks are automatically executed by ProductCart, including sending a Confirmation Email to the customer. Please consult the User Guide for more information on the exact list of tasks that are performed.<br /><br />If you leave this feature unchecked, the order remains 'Pending', regardless of whether payment has already been collected. This is the recommended method of processing orders as it allows you to verify the validity and accuracy of the order before processing it.<br /><br />Note that you should not finalize the payment portion of the order until products have shipped or services have been provided. This means that you should set your payment gateway (if any) to authorize the transaction, without capturing the funds (choose 'authorize' vs. 'sale' or 'capture')."
			
		case 302
			pcStrTitle = "Setting &amp; changing the payment status"
			pcStrDetails = "In ProductCart v3 and above the <strong>payment status</strong> is managed and reported separately from the <strong>processing status</strong>. For example, for whatever reason you might decide to process an order (send the confirmation email, ship the products, provide a license key for a software download, etc.) even if payment has not been collected for that order (e.g. a customer that pays you &quot;Net 30&quot;).<br /><br />This way you can track whether or not you have been paid for a certain order independently of whether the customer received the products or services that he/she ordered.<br /><br />ProductCart lets you specify which payment status should be assigned to an order by the system when the order is placed (*). Typically...<ul><li>The payment status is <strong>Pending</strong> when payment has not been processed (e.g. NET 30, offline credit card processing, or offline check processing, etc.). You will normally want to choose this option when you set up an offline payment option.<br /><br /></li><li>The status is <strong>Authorized</strong> when payment has been authorized, but funds have not been captured (e.g. online credit card processing with the payment gateway set to 'Authorize' vs. 'Sale' or 'Capture'). You will normally want to choose this option when you configure your payment gateway to authorize a transaction without capturing the funds.<br /><br /></li><li>The status should be <strong>Paid</strong> if payment has already been finalized by the time the order is placed.</li></ul><u>Note</u>: changing the payment status in the &quot;Payment Status&quot; tab of the &quot;Order Details&quot; page in the Control Panel is for administrative purposes only. That is: if a payment gateway was involved in the order, the payment gateway is <u>not</u> contacted when you change the payment status.<br><br>In some cases (e.g. when using the NetSource Commerce Payment Gateway), a separate section is displayed below the Update Payment Status area and allows you to carry out tasks related to the payment status, such as refunding an order.<br /><br /><hr>(*) Previous versions of ProductCart made an assumption on the payment status based on the payment settings: so an order was always considered 'processed' when payment was authorized and captured in real-time. You now have more flexibility in managing your orders as no assumptions are made. You can define how orders should be handled."
			
		case 303
			pcStrTitle = "Custom Payment Options: general information"
			pcStrDetails = "You can create an unlimited amount of special payment options on your store. What's common to all of them is that there is no real-time exchange of funds (e.g. Net 30, Purchase Order, Mail a Check, etc.).<br /><br /><strong>Option Name</strong>: the name of the payment option shown in the storefront (e.g. Pay by Check). It can be a maximum of 70 characters.<br /><br /><strong>Apply to Wholesale Customers Only</strong>: check this setting if you want this payment option to only be available to a wholesale customer (or a customer that belongs to a <a href='AdminCustomerCategory.asp' target='_blank'>Pricing Category</a> that has been setup to have wholesale privileges)."
			
		case 304
			pcStrTitle = "Custom Payment Options: processing fee"
			pcStrDetails = "You have the option of associating a processing fee with a payment option. For example, assume that processing a check that is sent to you by mail costs you more than processing a credit card transaction, and therefore you want to discourage customers from using that payment option. You can charge a processing fee to do so.<br /><br />The fee can be an absolute amount or a percentage of the order total. It is shown in the storefront in the drop-down that lists the available payment options."
			
		case 305
			pcStrTitle = "Pricing Categories: wholesale privileges"
			pcStrDetails = "If a customer is assigned to a pricing category that has Wholesale Customer Privileges, the customer is treated as a wholesale customer. This means that the customer will have access to wholesale-only payment options, wholesale discounts, etc."
			
		case 306
			pcStrTitle = "Pricing Categories: &quot;across the board&quot; vs. &quot;product by product&quot;"
			pcStrDetails = "<strong>Across the board</strong>: If you plan to give customers that are assigned to this category a special discount that applies to all products in your store 'across the board', choose this setting. This setting saves you time because you will not need to assign a special price to each product. ProductCart will calculate the special price &quot;across the board&quot;, based on the discount that you specify. Note that you can overwrite the price determined by this automatic calculation by manually entering a different price on a product by product basis.<br /><br />Technically speaking, in this first scenario the <i>pcCC_Pricing</i> table is not populated with any price unless you choose to overwrite the default price (online or wholesale price minus across the board discount), with a special price.<br /><br /><strong>Product by product</strong>: If prices vary product by product, choose this setting. This assumes that <u>for each product</u> in your store, you will specify the special price that customers that are assigned to this pricing category will pay.<br /><br />Technically speaking, in this scenario, the <i>pcCC_Pricing</i> table is populated with a price for each product in the store."
			
		case 307
			pcStrTitle = "Pricing Categories: assigning the first &quot;product by product&quot; price"
			pcStrDetails = "Using this setting you can assign to all products in your store a price based on the <strong>online price minus the discount that you enter</strong>.<br /><br />Assign a first &quot;product by product&quot; price for this pricing category by applying a discount off the current online price. You can then edit the price on a product by product basis when editing products or through the Global Changes feature.<br /><br />Since ProductCart will update prices for all products in your catalog, <strong>this task may take some time</strong>.<br /><br />If you leave this field blank, no prices will be assigned to your products with regard to the current pricing category. This means that all customers that are associated with this pricing category will be charged the regular online price. You can assign a price to your products for this pricing category at any time when adding/editing a product or through the Global Changes feature."
			
		case 308
			pcStrTitle = "Pricing Categories: overwriting the &quot;across the board&quot; price"
			pcStrDetails = "If you enter a price in this field, this new price will overwrite the default price calculated based on the &quot;across the board&quot; percentage that you specified when you created this pricing category. To view or edit pricing categories, <a href='AdminCustomerCategory.asp' target='_blank'>click here</a>." 
			
		case 309
			pcStrTitle = "Processing an order"
			pcStrDetails = "When you process an order, a number of things happen on your store. For a list of all the tasks that ProductCart performs when an order is processed, please refer to the ProductCart User Guide. You can process orders one by one or batch process multiple orders at once using the &quot;Batch Process Orders&quot; feature. Please note:<ul><li>When you <u>individually process an order</u>, your payment gateway (if any is being used to collect payment) is NOT contacted to capture a previously authorized transaction. You will have to perform that task separately by logging into the payment gateway's administration area.<br /><br /></li><li>If you <u>batch process</u> multiple orders at once and you are using either <strong>Authorize.Net</strong>, <strong>LinkPoint API</strong>, <strong>NetBilling</strong>, <strong>USAePay</strong>, <strong>PayPal PayFlow Pro</strong> or <strong>PayPal Website Payments Pro</strong> as your payment gateway, then the payment gateway is contacted by ProductCart and the orders - previously authorized - are captured and will be settled. This is only true if you are using one of these 4 payment gateways. In all other cases the payment gateway is <u>not</u> contacted to finalize payment.</li></ul>"
			
		case 310
			pcStrTitle = "Cancelling an order"
			pcStrDetails = "When you cancel an order, a number of things happen on your store. For a list of all the tasks that ProductCart performs when an order is cancelled, please refer to the ProductCart User Guide.<br /><br />Please note that when you cancel an order your payment gateway (if any is being used to collect payment) is NOT contacted to automatically void or credit the corresponding transaction a previously authorized transaction. You will have to perform that task separately by logging into the payment gateway's administration area."
			
		case 311
			pcStrTitle = "Resetting the order status"
			pcStrDetails = "You can reset the order status to change the way it is recorded in the store database. However, you should rarely use this feature. This is because &quot;resetting&quot; an order's processing status is not like &quot;updating&quot; its status. When you &quot;reset&quot; the status of an order, you are simply changing the way the order is saved into the store database, without altering anything else. For example, inventory levels are not adjusted, &quot;Reward Points&quot; are not adjusted, no notification e-mails are sent, etc."
			
		case 312
			pcStrTitle = "Product prices on the Edit Order page"
			pcStrDetails = "The Edit Order page attempts to find a balance between automatically applying price calculations based on store settings (e.g. quantity discounts) and allowing the Store Manager to override such settings to 'force' a price (e.g. changing the unit price manually to a special value for a certain order). Please note the following about how product prices are handled on the Edit Order page: <ul><li>The <b>Unit Price</b> is calculated as follows: Product’s Base Price – Pricing Category Discount (<i>if any</i>) - Quantity Discounts (<i>if any</i>).<div style='padding-top: 4px;'>If the product is a Build To Order product, the Unit Price also includes the prices of all items that were selected by the customer when configuring the BTO product. However, it does not include 'Additional Charges', since they are independent of the quantity purchased. Additional charges are included in the 'Price' (row total).</div></li><li>The <b>Price</b> is normally the Unit Price times the number of units purchased. Option prices are not included. They are shown right below. If there are quantity discounts that apply to the option prices, the prices shown reflect those discounts. When the product is a Build To Order product, the price is the price of the fully configured BTO product, including everything (selections, discounts, additional charges).</li><li><b>Quantity Discounts</b> are reflected in the displayed Unit Price.</li><li>Any price differential associated with <b>Product Options</b> is <u>not</u> reflected in the displayed Unit Price, but rather shown separately. </li><li><b>Updating prices</b>. You may manually change the Unit Price to any monetary value (with up to 2 decimals). When you click 'Update', the following is true:<ul><li>Quantity Discounts <u>will not</u> be applied. The value you enter is the price used in the calculation (minus any option price differential, as mentioned above).</li><li>Category-based Quantity Discounts <u>will</u> be applied and re-calculated based off the new Unit Price.</li></ul>Clicking 'Update' a second time after modifying the Unit Price will assume you want to recalculate prices considering all variables, and therefore Quantity Discounts are now applied. So if you intend to manually overwrite the Unit Price, 'overriding' account quantity discounts, setting the Unit Price should be the last task performed on the edit order screen.</li><li>If the customer belongs to a <b>Pricing Category</b>, note that the special pricing associated with that pricing category is lost when you 'Restore' the product price as the system looks up the default online price for the product.</li><li><strong>Category-based Quantity Discounts</strong> affect the unit price of a product, but are applied to the entire order as a discount, not to the an individual product sub-total calculation. This is because by definition they are not product-specific, but rather are calculated based on the total quantity of purchased products belonging to a certain category. So you should look for the discount value in the Discount section of the page.</li><li>If you <strong>add new items</strong> to an order (e.g. new products, new options, new coupons, etc.) you must click the 'Update' button. Do so <u>after</u> you have added the new information. Otherwise ProductCart will not recalculate quantity discounts on the prices.</li></ul><u>Technical Note</u>: the unit price saved into the 'ProductsOrdered' table ('unitPrice' field) is different from the one displayed on this page; it includes Options Prices."
			
		case 313
			pcStrTitle = "About 'restoring' the Unit Price"
			pcStrDetails = "When you 'Restore' the Unit Price, ProductCart will copy the product's current 'Online Price' into the field. No change is applied to the order until you 'Update' it.<br><br>Please note: <ul><li>If the product is a BTO product, the base price is restored, not the default price (base price + default options). If you need to change the BTO configuration, use the 'Edit Configuration' link. This means that - in most cases - the 'Restore' feature should not be used with BTO products.</li><li>If the customer belong to a <u>Pricing Category</u>, the special pricing associated with that pricing category is lost when you 'Restore' the product price.</li></ul>."
			
		case 314
			pcStrTitle = "Editing an order: shipping and taxes"
			pcStrDetails = "When you edit an order, shipping and taxes are not automatically recalculated. Specifically: <ul><li><b>Shipping Rates</b>. Since the total order weight can change when you edit an order, you will need to look up again the shipping rate associated with the order. This is not done automatically when you update the order as it often involves connecting to UPS, FedEx or another shipping provider. ProductCart will recalculate all available shipping options for the order when you click on 'Check Real-Time Rates'.</li><li><strong>Taxes</strong>. Unless there is a value in the 'Calculate' input field, taxes will not be recalculated. Use the first input field for a percentage rate or the second one for a flat amount. Use the 'Check tax rates' link to recalculate taxes based on the new order information. Check the 'Tax' checkbox next to each taxable item on the 'Edit Order' page, then click on 'Update' to recalculate taxes on the order.</li></ul>."
			
		case 315
			pcStrTitle = "Google Analytics - Refund &amp; Cancellations: order information"
			pcStrDetails = "When an ecommerce transaction is posted to Google Analytics, it is made up of two parts: one containing general order information, and one containing item information (the products that were purchased). So when you post a refund/cancellation, you need to include both.<br><br>Your store will create and post the refund/cancellation transaction based on the information you specify on this page.<br><br>(1) First, specify the '<strong>General Information</strong>': the total amount to refund, any shipping charges to refund, and any tax amount to refund (taxes and shipping are part of the total amount: they are separated out only for reporting purposes).<br><br>(2) Then, specify which items are being returned under '<strong>Item Information</strong>'. <u>If you don't want to return an item (partial return), set its quantity to 0</u>. When the 'Returned/Refunded' quantity is 0, that means that the item is not being returned/cancelled. This should be reflected in a refund amount that is less than the orginal order total in the 'Order Information' section."
			
		case 316
			pcStrTitle = "Batch Shipping Orders"
			pcStrDetails = "Batch shipping orders allows you to update the order status of multiple, processed orders to 'Shipped'. You can send an e-mail to customers while updating the order status. You can also import a list of orders to perfom this task in a semi-automatic way: this is useful - for example - if you need to import a substantial number of orders with shipment tracking information.<br><br>Note that this feature has some <strong>limitations</strong>:<ul><li>You <u>cannot print labels</u> (e.g. UPS, FedEx, USPS) when using the batch shipping feature. You need to use the Shipping Wizard individually for each order.</li><li><u>Not all processed orders can be batch-shipped</u>. Specifically, you cannot batch shipping orders that contains one or more products that are currently out-of-stock, back-ordered, or drop-shipped. In the case of out-of-stock and back-ordered items the order is shown on the page, but a message indicates that it cannot be batch-shipped. Orders containing drop-shipping items are not shown.</li><li>Unlike when using the Shipping Wizard, you <u>cannot personalize the e-mail message</u> received by each customer. The default template will be used for all customers whose orders are batch-shipped (<a href='emailsettings.asp' target='_blank' onClick='JavaScript: window.close()'>edit it</a>), with the information specific to their orders (i.e. shipping method, shipping date, and tracking number, if any).</li></ul>"
			
		case 317
			pcStrTitle = "Guest Checkout: Order Code"
			pcStrDetails = "When customers checkout as guests (i.e. without saving a password and without logging into an existing account), ProductCart automatically generates an order code that they can use, together with their e-mail address, to log in and review their order status. If they contact you about their order status, remind them that they can use the <strong>Order Code</strong> to log into the <a href='../pc/Checkout.asp?cmode=1' target='_blank'>customer service area</a> of the storefront and do so."
			
		case 318
			pcStrTitle = "Pricing Categories: Override Not For Sale setting"
			pcStrDetails = "This option allows you to sell to customers that belong to the selected Pricing Category products that are &quot;Not For Sale&quot; for other customers. This can be useful when creating <strong>Private Sales</strong> or a <strong>Private Shopping Club</strong>: only customers that have been assigned to the selected Pricing Category will be able to purchase the products (and see the special price associated with that Pricing Category, if different from the regular prices)."
			
			
		' General administration tools
		
		case 400 ' Using the built-in HTML editor
			pcStrTitle = "Using the built-in HTML editor"
			pcStrDetails = "ProductCart includes an advanced, browser-based HTML editor published by <a href='http://www.innovastudio.com/' target='_blank'>Innova Studio</a>.<br /><br />To format your text, you can either paste your own HTML code into the HTML editor window or use the editing tools that are provided.<br /><br />If you decide to paste your own HTML code (e.g. from an application just as Macromedia Dreamweaver), make sure to include only code that is located after the opening BODY tag and before the closing BODY tag." 
			
		case 401 ' Using short product description
			pcStrTitle = "Using the short product description"
			pcStrDetails = "When you add a short description to a product, the long description is moved to a separate area of the page, and only the short description is shown at the top. A 'More details...' link (which you can edit using the 'includes/languages.asp' file) will allow customers to view the long description.<br /><br />We recommend that you don't use HTML code in the Short Description, except for simple HTML tags such as making text bold or underlined. This is because the short description is used in several areas of the store beyond the product details page (e.g. search results pages, browse by category pages, etc., when the products are shown in a list)." 
							
		case 402 ' Global changes: recalculating prices
			pcStrTitle = "Setting prices based on another price or cost"
			pcStrDetails = "This feature allows you to quickly update product prices based on a calculation that uses another price or the product cost as the starting point.<br /><br />For example, if the current online price for a certain product 50 and you select the &quot;list price&quot; from the first drop down, then enter the percentage change of &quot;125&quot; and select the &quot;online price&quot; from the second drop down, the list price is recalculated to: 62.50.<br /><br />This feature is very useful to set prices based on a margin on the product cost. For example, assume that you want to have a 30% margin on all &quot;computer keyboards&quot; on your store: you quickly do this by setting the online price as 130 percent of the cost." 
			
		case 403
			pcStrTitle = "Quick price changes: how it works"
			pcStrDetails = "The price change can be either positive or negative, % or absolute amount (e.g. for a 8% decrease in price, enter &quot;-8&quot; and select &quot;% change&quot;)." 
			
		case 404
			pcStrTitle = "E-mail settings: user registration notification"
			pcStrDetails = "The store manager always receives an email when an order is submitted. Here you can turn on or off a notification email that the system sends when a customer <u>registers</u> with the store, regardless of whether the customer places an order.<br /><br />A welcome e-mail is also sent to the customer upon registration.<br /><br /><u>Technical Note</u>: You can edit the text used in that welcome e-mail by editing the file &quot;includes/languages.asp&quot;. Look for the text strings &quot;storeEmail_20&quot; to &quot;storeEmail_23&quot;." 
			
		case 405
			pcStrTitle = "E-mail settings: Order Received e-mail message"
			pcStrDetails = "This text will be included in the e-mail that is automatically sent to customers after an order has been placed. It should not say that the order is confirmed or processed, but only that it has been received. The confirmation e-mail is sent separately, once the order has been processed." 
			
		case 406
			pcStrTitle = "E-mail settings: Order Confirmation e-mail message"
			pcStrDetails = "This text will be added to the confirmation e-mail that is automatically sent to customers after an order has been processed. It will be displayed before the order details in the body of the e-mail message.<br /><br />This text is always included in the confirmation e-mail regardless of whether orders are processed manually via the Control Panel or automatically (e.g. real-time credit card processing and order processing).<br /><br />If an order is processed manually through the Control Panel, then you also have the ability to add some additional, order-specific comments. Refer to the User Guide for more information about processing orders."
			
		case 407
			pcStrTitle = "E-mail settings: Order Confirmation e-mail message"
			pcStrDetails = "This text will be added to the e-mail that is automatically sent to customers to notify them that their order has been shipped. It will be displayed before the shipping details in the body of the e-mail message.<br /><br />When you ship an order through the Control Panel, you also have the ability to add some additional, order-specific comments. Refer to the User Guide for more information about shipping orders."
			
		case 408
			pcStrTitle = "E-mail settings: Order Confirmation e-mail message"
			pcStrDetails = "This text will be added to the e-mail that is automatically sent to customers to notify them that their order has been cancelled. You also have the ability to add some additional, order-specific comments at the time you cancel the order."
			
		case 410
			pcStrTitle = "Store settings: hiding the category drop-down on the search page"
			pcStrDetails = "If the <a href='../pc/search.asp' target='_blank'>advanced search page</a> takes a while to load, you can speed it up by either disabling the category drop-down completely, or limiting to top-level categories only.<br /><br />Especially on stores that have several hundred categories, using these options can substiantially improve page loading time for the advanced search page."
			
		case 411
			pcStrTitle = "Store settings: allow the purchase of out-of-stock products"
			pcStrDetails = "Use this setting to allow or disallow the purchase of out-of-stock products store-wide. If inventory is an issue on your store, you will typically disallow the purchase of out of stock items using this setting. On a product by product basis, you can then override this setting so that specific products can be purchased even if they are out of stock."  
			
		case 412
			pcStrTitle = "Store settings: turning your store ON or OFF"
			pcStrDetails = "Use this feature to temporarily turn on/off your store (e.g. you are doing maintenance). You may also edit the message shown to your customers (you can use HTML tags).<br /><br />It is definitely recommended that you turn your store off before updating or upgrading your copy of ProductCart to a newer version." 

		case 413
			pcStrTitle = "Store settings: turn the Wish List on and off"
			pcStrDetails = "When this option is set to on, your customers will be able to save products to their accounts so that they may review them and purchase them at a later time." 
			
		case 414
			pcStrTitle = "Store settings: turn the Tell A Friend feature on and off"
			pcStrDetails = "When this option is enabled;, a &quot;Tell a Friend&quot; button will be added to your product details pages, allowing store visitors to e-mail a link to that page to a friend." 
			
		case 415
			pcStrTitle = "Store settings: turning your store into a catalog-only"
			pcStrDetails = "This store setting allows you to turn off your customers ability to place an order. You can do so for all customers or just for retail customers. When you turn off the ability to place an order for all customers, your store will behave as an online catalog instead of an online store. Customers will be able to browse and search the catalog, but not place orders."
			
		case 416
			pcStrTitle = "Store settings: Maximum number of Products"
			pcStrDetails = "This is the <u>maximum number of different products</u> your customers will be able to purchase at one time, regardless of the quantity ordered for each product. In order to limit the amount of server resources used for a user session, the number is structurally limited to 100 (i.e. 100 different products added to the shopping cart). Consult the User Guide for more information."
			
		case 417
			pcStrTitle = "Store settings: Maximum number of units per product"
			pcStrDetails = "This is the <u>maximum quantity</u> your customers will be able to order for any product in your catalog (i.e. 20 means that they can only order up to 20 units of a certain product)."
			
		case 418
			pcStrTitle = "Store settings: Hide discount code (electronic coupon) input field"
			pcStrDetails = "By default this option is set to &quot;No&quot;, and an input field is shown at the bottom of the Payment Panel during the one page checkout process (pc/OnePageCheckout.asp). If there are no discounts or gift certificates available in your store, you can hide the field by setting this option to &quot;Yes&quot;."
			
		case 419
			pcStrTitle = "Store settings: category, product, and search preview"
			pcStrDetails = "ProductCart v4 contains advanced interface elements that take advantage of a set of technologies that have been defined as AJAX (<a href='http://en.wikipedia.org/wiki/AJAX' target='_blank'>more information on AJAX</a>).<br /><br />Among the AJAX-enhanced components you will find:<ul><li><strong>search preview</strong> feature, which tells you how many results will be returned before you perform the search</li><li><strong>product details preview</strong> feature, which provides product details when you mouse over a product thumbnail, without visiting the product details page</li><li><strong>category content preview</strong> feature (<em>new in ProductCart v4</em>), which provides information on the subcategories and products contained in that category</li></ul>In ProductCart v4 and above, you can separately turn on and off the different <em>previews</em>. In previous versions, you could only turn on and off the entire feature.<br><br><strong>Note on Performance</strong><br>Although we are not aware of any performance issues linked to this feature, we would still like to mention that this feature increases the number of &quot;calls&quot; to the store database. Therefore, it is strongly recommended that you only use this feature when your store is powered by a MS SQL database. In addition, you should try disabling this feature if you are experiencing performance issues: compare the store performance when the feature is turned ON vs. OFF to understand whether this feature is negatively impacting performance on your store."
			
		case 420
			pcStrTitle = "Store settings: Turn error handler ON/OFF (debugging tool)"
			pcStrDetails = "This setting should always be set to ON except for the case in which a technician is troubleshooting a problem with the software. When set to ON, the error handler that is built into ProductCart always returns a friendly error in the storefront (this is also a security precaution). When set to OFF, it returns the actual error that has occurred."
			
		case 421
			pcStrTitle = "Store settings: Disallow Return Merchandise Authorization request"
			pcStrDetails = "Your store allows customers to request an authorization to return a product that they have previously purchased (RMA = Return Merchandise Authorization). The store manager is notified whenever a request is submitted. If you don't want customers to be able to request an RMA, you can disable this feature here."
			
		case 422
			pcStrTitle = "Store settings: Turn built-in Help Desk ON/OFF"
			pcStrDetails = "When customers log into their account and view previous orders, they can communicate with you using the Help Desk that is built into ProductCart. If you don't want customers to have access to this feature, you can turn it off using this setting. For more information on the Help Desk, please see the User Guide."

		case 423
			pcStrTitle = "Store settings: Allow checkout without registering a password"
			pcStrDetails = "When this feature is set to YES, customers will be able to checkout without having to enter a password.<br /><br />Note that even when this feature is on, a customer account <u>is</u> created as an order cannot be saved to the database without saving customer information. A random password is added to the customer account and the customer will be able to retrieve it, if they ever wish to do so (e.g. to log in and check the status of an order or track a package)."
			
		case 424
			pcStrTitle = "Store settings: Product details page display settings"
			pcStrDetails = "You can choose among 3 different display settings for the product details page:<ol><li>Two-column layout with the product image on the left</li><li>Two-column layout with the product image on the right</li><li>One-column layout with the image in the middle (ideal for products that have wide images, or for products that don't have any image)</li></ol>You can override the store-wide setting at the category level (i.e. all products within the same category), and even at the product level (i.e. product specific setting).<br /><br />Here is a graphical example of how the different layout settings affect the same product page.<br /><img src='images/pcv3_prdDetailsPage.gif' alt='Product details page alternative layouts' vspace=10 /><br /><br /><u>Technical Note</u>: The product details page selects the style for the page in the following order of priority (the variable on viewPrd.asp is called &quot;pcv_strViewPrdStyle&quot;):<ul><li>querystring (e.g. ViewPrdStyle=o)</li><li>product-specific setting</li><li>category-specific setting</li><li>storewide setting</li></ul>"
			
		case 425
			pcStrTitle = "Managing Products: multiple product images"
			pcStrDetails = "You can display multiple product images on the product details page (pc/viewPrd.asp). The additional images are shown as follows:<ul><li>A thumbnail for each additional image is shown below the main product image. 3 thumbnails are shown per row (you can change this value by editing it directly on the product details page, pc/viewPrd.asp). The size of the thumbnail is determined automatically based on how many thumbnails are displayed and how wide the main product image is.</li><li>The General size image is shown on the product details page in place of the standard product image when you mouse over the corresponding thumbnail.</li><li>If you click on any of the thumbnails, a new window will be displayed, where the Detail View image is shown.</li></ul>"
			
		case 426
			pcStrTitle = "Shipping Settings: default shipping provider"
			pcStrDetails = "When more than one shipping provider is active in a store, the customer will be shown a drop-down on the shipping rates selection page. Here you can specify which shipping rates will be loaded first on the page.<br /><br />Please note that ProductCart shows only one shipping provider on the shipping selection page (e.g. UPS <u>or</u> FedEx rates, not both at the same time) because of requirements set by the shipping providers themselves."
			
		case 427
			pcStrTitle = "Store settings: category display settings"
			pcStrDetails = "A category page can contain both products and sub-categories. You can specify different display settings for different categories, and display sub-categories differently from the way products are displayed within the category. For example, you could organize sub-categories horizontally, but display products vertically within the same category, or viceversa.<br /><br />You can choose among 3 different display settings for the way categories and subcategories are displayed in the &quot;Browse by Category&quot; pages. These settings can be set store-wide using the &quot;Store Settings&quot; page or at the category level (overriding the store-wide setting for sub-categories within the selected category). The settings are:<ol><li>As a list (category names only)</li><li>With category names and thumbnail images</li><li>In a drop-down</li></ol>Here is a graphical example of how a category page that contains sub-categories is organized.<br /><img src='images/pcv3_display_categories.gif' alt='Example of a category page' vspace=10 />"
			
		case 428
			pcStrTitle = "Store settings: number of categories per page"
			pcStrDetails = "When categories are <u>not</u> displayed in a drop-down, you can set the number of categories to be shown in each row (e.g. 3) and the number of rows per page (e.g. 4). This sets the total number of categories shown on the page (e.g. 12 in this example). If there are more categories than the number specified here, page navigation is shown."
			
		case 429
			pcStrTitle = "Store settings: product display settings"
			pcStrDetails = "You can display products in your storefront in four different ways. These settings apply to all areas of the storefront (e.g. categories that contain products, search results pages, specials, new arrivals, best sellers, etc.) unless they are overwritten for specific areas (e.g. &quot;specials&quot;) or specific categories (<a href=""manageCategories.asp"" target=""_blank"">manage categories &gt;&gt;</a>).<br /><br />You can choose among four display settings for how products are presented to your customers:<ol><li>Display items horizontally (large product thumbnail)</li><li>Display items vertically</li><li>Display items vertically (list view - small product thumbail)</li><li>Display items vertically (list view - small product thumbnail) with the ability to add multiple products to the cart in one step.</li></ol>Here is a graphical example of how the different layout settings affect the same page.<br /><img src='images/pcv3_display_products.gif' alt='How products can be presented in the storefront' vspace=10 /><br /><br /><u>Technical Notes</u>: Every page that shows products in the storefront selects the style for the page in the following order of priority (the variable is called &quot;pageStyle&quot;):<ul><li>querystring (e.g. pageStyle=m)</li><li>category settings (or other page settings, e.g. &quot;Manage Best Sellers&quot;)</li><li>storewide setting</li></ul>The variable values are:<ul><li>Display items horizontally = h</li><li>Display items vertically = p</li><li>List view = l</li><li>List view + multiple add to cart = m</li></ul>"
			
		case 430
			pcStrTitle = "Store settings: number of products per page"
			pcStrDetails = "The combination of <u>products per row</u> and <u>rows per page</u> defines how many products are shown on each page. If there are more products than the number specified here, page navigation is shown. You can override these store-wide settings at the category level.<br /><br /><ul><li><b>Products per Row</b>: When products are shown <u>horizontally</u>, you can set the number of products to be shown in each row (e.g. 3) and the number of rows per page (e.g. 4). This sets the total number of products shown on the page (e.g. 12 in this example). This setting does not apply to products shown vertically or in a list.<br /><br /></li><li><b>Rows per Page</b>: This setting applies to all display types. When products are displayed vertically or in a list, it specifies the number of products shown on the page.</li></ul>"
			
		case 431
			pcStrTitle = "Store settings: show/hide SKU and small image"
			pcStrDetails = "When products are displayed in a list, you can specify whether you want the product part number (SKU) and a small product thumbnail to be displayed or not.<br /><br />Technical Note: the product thumbnail used when products are shown in a list is a smaller version of the thumbnail image specified for the product (if any). The size of the thumbnail is controlled by the pcStorefront.css cascading style sheet (&quot;pcShowProductImageL&quot; and &pcShowProductImageM&quot; classes)."
			
		case 432
			pcStrTitle = "Store settings: default product sorting preference"
			pcStrDetails = "When products are displayed within a category, they are sorted according to the preference specified here. You can override this store-wide preference at the category level by manually set the order in which products should be displayed within the category. To do this, look for a category, select &quot;Products&quot; to view the products listed within the category, and sort the products using the corresponding input field."
			
		case 433
			pcStrTitle = "Store settings: user-selected product sorting"
			pcStrDetails = "Unless you disable this feature, visitors to your store will be able to re-order products shown on a page. A drop-down allows them to do so by name, price or part number. This feature is typically disabled only on store that contain a very limited amount of products, in which case showing the drop-down might not be necessary."
			
		case 434
			pcStrTitle = "Store settings: allow users to nick-name orders"
			pcStrDetails = "When this feature is active an input field is shown during checkout, which allows users to name the order that they are placing. Once the order has been placed, they can easily rename it.<br /><br />This feature is useful on stores where customers place many <b>repeat orders</b> (e.g. office supplies), as it allows them to quickly locate a previous order and repeat it using the &quot;Repeat Order&quot; feature (this feature makes placing an order similar to a previous one very easy and fast)."
			
		case 435
			pcStrTitle = "Store settings: customizing buttons and icons"
			pcStrDetails = "You can use your favorite graphic program to create custom buttons and icons to be uploaded to your store. However, if you are familiar with <u>Adobe (Macromedia) Fireworks</u>, ProductCart includes the <strong>editable PNG files</strong> that were used to created these default buttons. Editing these files, rather than starting from scratch, can save you a great deal of time.<br /><br />The PNG files are located in the &quot;pc/images/sample&quot; folder."
			
		case 436
			pcStrTitle = "Store settings: managing content pages"
			pcStrDetails = "ProductCart includes a basic content management system that allows you to create and manage any number of Web pages, without using an external HTML editor.<br /><br />Click on <strong>Add New Content Page</strong> to add a new page. Use the built-in HTML editor to create the page or copy text or HTML code from another program. ProductCart uses an advanced HTML editor published by InnovaStudio. To learn how to best take advantage of this powerful tool, see the tutorials listed on: <a href='http://www.innovastudio.com/editor_tutorial.asp' target='_blank'>www.innovastudio.com/editor_tutorial.asp</a>.<br /><br /><ul><li>Check the <strong>Active</strong> option unless you want this page to remain inactive. Inactive pages cannot be accessed by your customers.<br /><br /></li><li>Check the <strong>Include store header &amp; footer option</strong> if you would like ProductCart to create a page that includes the store’s graphical interface. If this is the case, and if you decide to copy HTML code that you have created in another HTML editor (e.g. MS FrontPage, Macromedia Dreamweaver, etc.), make sure to include only code that is in between the &lt;body&gt; and &lt;/body&gt; tag. For example, you can certainly copy and paste an HTML table that you have created in your favorite HTML editor, but you should not copy an entire HTML page. If you wish to copy an entire HTML page into the Page Description field, then make sure not to select the Include store header &amp; footer feature.<br /><br /></li><li>Enter a <strong>Page Title</strong>: this is the title that is used by ProductCart’s header.asp file for the &lt;title&gt; tag for the page, assuming that you are using the dynamic meta tag generator mentioned in the Making Your Store More Search Engine Friendly section. Therefore, the title will not be shown in the body on the page.<br /><br /></li><li>Click on <strong>Add Content Page</strong> to save your new content page to the ProductCart database. All of the content pages that you have created are listed on the Manage Content Pages page. At any time you can edit a page using ProductCart’s built-in HTML editor.</li></ul>To allow customers to find the content pages that you have created, link to them from other pages on your Web site, and/or from your store’s navigation. You can copy the location of the page to your clipboard by using the button situated on the right side of the page URL.<br /><br /><u>Technical Note</u>: <strong>dynamically loading a list of content pages</strong>. The default version of header.asp contains some ASP code that dynamically loads content page titles and corresponding links from the store database. This allows you to create a list of links to these pages that is automatically updated every time you add a new page and remove/edit an existing page.<br /><br />You could place this code anywhere in your custom version of header.asp (or another file set up to query the ProductCart database) to load such list. If you copy and paste the code below, make sure to review and edit it in your HTML editor as it might contain extra line breaks that should not be there. Alternative, you can copy the same code directly from the default version of header.asp.<br /><br />&lt;% ' Show List of Content pages<br />&nbsp;&nbsp;sdquery=&quot;SELECT pcCont_IDPage,pcCont_PageName FROM pcContents where pcCont_InActive=0;&quot;<br />&nbsp;&nbsp;set rsSideCatObj=conlayout.execute(sdquery)<br />&nbsp;&nbsp;do while not rsSideCatObj.eof<br />%&gt;<br /><br />&lt;a href=&quot;viewContent.asp?idpage=&lt;%=rsSideCatObj(&quot;pcCont_IDPage&quot;)%&gt;&quot;&gt;&lt;%=rsSideCatObj(&quot;pcCont_PageName&quot;)%&gt;&lt;/a&gt;<br /><br />&lt;%<br />&nbsp;&nbsp;rsSideCatObj.MoveNext<br />&nbsp;&nbsp;loop<br />&nbsp;&nbsp;set rsSideCatObj=nothing<br />%&gt;<br />"
			
		case 437
			pcStrTitle = "Store settings: countries, states and provinces"
			pcStrDetails = "Your store comes with a list of countries that follow ISO standard names and codes (English language). For a few of those countries (Australia, Canada, United States and New Zealand) state/provinces have also been added to the system. They also follow ISO standards.<br /><br />You can add/edit the list of countries and/or states and provices at any time. If you do so, make sure to adhere to the <a href='http://www.iso.ch/iso/en/prods-services/iso3166ma/02iso-3166-code-lists/list-en1.html' target='_blank'>ISO standards</a> to ensure proper calculation of shipping charges via one or more of the shipping providers that your store has been integrated with.<br /><br />Please note:<ul><li><strong>State/Province Drop-Down Menu</strong>: If a country has one or more states/provinces associated with it, the &quot;state/province&quot; input field will automatically change to a drop-down including those states/provinces. This is the reason why the &quot;Country&quot; drop-down is located before the &quot;State/Province&quot; input field (or drop-down) whenever an address is shown.<br /><br /></li><li><strong>State/Province Input Field</strong>: If a country does not have any states/provinces associated with it, a text input field is shown instead of a drop-down.</li></ul>You can associate a state/province with a country when you add it or edit it.<br /><br />You can restore the default list of countries, states and provinces at any time by using the &quot;Restore Default&quot; button on the corresponding page."
			
		case 438
			pcStrTitle = "Managing Products: export for re-import (Reverse Import Wizard)"
			pcStrDetails = "This feature allows you to export product information so that it can easily be re-imported into ProductCart. This way you can easily export fields that you need to update, edit them, and then import them back into your store using the Import Wizard.<br /><br />Product information is exported to a CSV file. If you open it in MS Excel and plan to import it back into your store as an Excel file, make sure to set the &quot;IMPORT&quot; data range and follow the other instructions <a href='index_import_help.asp' target='_blank'>listed here</a>."
			
		case 439
			pcStrTitle = "Managing Categories: category images"
			pcStrDetails = "The <strong>small category image</strong> is shown on the pages where the customer is &quot;browsing by category&quot;, if the store has been set up to use category images (if you set up your store to only list the category names or list the categories in a drop-down, the small category image is not used).<br /><br />The <strong>large category image</strong> (if any) is used when the storefront is displaying products and/or sub-categories within the category. It is shown at the top of the page, above the long category description."
			
		case 440
			pcStrTitle = "Managing Categories: category descriptions"
			pcStrDetails = "The <strong>short category description</strong> is shown on the pages where the customer is &quot;browsing by category&quot;, if the store has been set up to show one category per row. If the store has been setup to show sub-categories in a drop-down, or with more than one per row, the short category description is not used.<br /><br />The <strong>long category description</strong> is used when the storefront is displaying products and/or sub-categories within the category. It is shown at the top of the page, below the long category description."
			
		case 441
			pcStrTitle = "Reports: conversion rate"
			pcStrDetails = "The conversion rate tells you how many of the new customers that registered with your store actually placed an order. Please note that if you upgraded to ProductCart from a previous version, this report will only be meaningful for a date range starting after the upgrade date. This is because before v3 ProductCart did not record the date in which a customer account was created, and therefore there was no way for the software to determine whether a customer account was created within the date range that you indicated."
						
		case 442
			pcStrTitle = "PayPal Action Buttons"
			pcStrDetails = "<p><u><strong>Void</strong></u><br /> This PayPal action is available for the payment statuses authorized and paid. When you &quot;Void&quot;, the payment status becomes voided and order status is canceled.</p><p><u><strong>Capture</strong></u><br /> This PayPal action is available for the payment status authorized. When you &quot;Capture&quot;, the payment status becomes paid and order status is processed.</p><p><u><strong>Refund</strong></u><br /> This PayPal action is available for the payment status paid. When you &quot;Refund&quot;, the payment status becomes refunded and order status remains as is.</p><p><u><strong>Reauthorize</strong></u><br /> This PayPal action is available for a transaction expires. A transaction expires 29 days after the first authorization. When you &quot;Reauthorize&quot;, the payment status becomes authorized and order status is not changed (It should be pending).</p>"	
			
		case 443
			pcStrTitle = "Printing Invoices, Packing Slips and Pick Lists"
			pcStrDetails = "<p>You can easily print invoices, packing slips or a &quot;pick list&quot; for the orders you select.<ul><li><strong>Invoice</strong>: a printer-friendly version of the order. It includes all order details.</li><li><strong>Packing Slip</strong>: a printer-friendly version of the order meant to be included in a shipment. It does not contain any prices.</li><li><strong>Pick List</strong>: a simple list of products associated with the selected orders. This can be useful as a hand-out for picking products from your warehouse when you are preparing your shipments.</li></ul></p>"	
			
		case 444
			pcStrTitle = "Generate Navigation: styling the links with CSS"
			pcStrDetails = "<p>By using cascading style sheets (CSS) you can change the way the link is displayed. For example, you can set the font color, size, whether it should be underlined or not, etc. Make sure to add the class that defines the style to a stylesheet that is loaded by <em>pc/header.asp</em>. For example, you could add a new class to 'pc/pcStorefront.css'.</p><p>If you are generating <strong>table rows and cells</strong>, the CSS class is assigned to all the table cells generated by the system. If you are using an <strong>unordered list</strong>, you can assign both an ID and a class to the top-level &lt;ul&gt; tag, which allows you to then manipulate the list with JavaScript (e.g. <a href=http://labs.adobe.com/technologies/spry/articles/menu_bar/index.html target=_blank>SPRY navigation</a>). You can then style all of the other tags in the unordered list using <a href=http://www.google.com/search?q=descendant+selectors+css target=_blank>descendant selectors</a> in your CSS document.<p>If you are interested, the Internet is full of information about Cascading Style Sheets. To get started, go to any search engine and do a search for 'css'.</p><p>For more information, please <a href=http://wiki.earlyimpact.com/how_to/add_category_navigation target=_blank>see the ProductCart WIKI</a> on this topic.</p>"	
			
		case 445
			pcStrTitle = "Generate Navigation: relative vs. absolute links"
			pcStrDetails = "<p>You can choose whether to create a navigation tree that uses relative vs. absolute links. Relative links simply point to the files in the same folder (the 'pc' folder). Absolute links use the full URL to the page (e.g. 'http://www.mystore.com/shop/pc/...').</p><p>Some search engine optimization experts say that your link structure should be consistent across your Web site. So if you are using absolute links elsewhere on your Web site, you should use absolute links in the store navigation as well.</p>"	
			
		case 446
			pcStrTitle = "Using images with products and categories"
			pcStrDetails = "<p><strong>What to enter in the field</strong><br>For each image field, enter the file name, not the path or URL to the image. For example &quot;myImage.gif&quot;. If you are using the &quot;Upload and Resize&quot; feature, the image file names are entered automatically for you.</p><p><strong>Uploading Images</strong><br>All images are located in the &quot;pc/catalog&quot; folder. You can upload an image to that folder directly using ProductCart. If your Web server supports it, you can use the handy &quot;<u>Upload and Resize</u>&quot; feature. See the User Guide for more details.</p><p><strong>Storing Images in Sub-Folders</strong><br>If your store has a large number of image files, you can organize them in sub-folders of the Catalog folder. Do this using your FTP program. Then, when entering your file name, add the name of the sub-folder. For example, if you add the sub-folder &quot;large-images&quot; to the &quot;catalog&quot; folder and FTP the image &quot;myLargeImage.gif&quot; to that folder, you will enter &quot;large-images/myLargeImage.gif&quot; in the image field.</p>"
			
		case 447
			pcStrTitle = "Back-ordering settings"
			pcStrDetails = "<p>When you enable back-ordering for a product, the product remains available for sale even if it is out of stock. This is different from the 'Disregard Stock' feature in that the system will track inventory. Inventory is not ignored, but the sale of the item is allowed even when out of stock (e.g. products that are regularly re-stocked).</p><p>Customers are told that the product is not immediately available through a message that indicates that the item typically ships in a certain number of days. You can specify the number of days through the corresponding field.</p>"	
			
		case 448
			pcStrTitle = "Low stock notification"
			pcStrDetails = "<p>When you activate this feature, the store manager receives an e-mail whenever a product's stock level falls below the number of units entered in the corresponding field.</p>"	
			
		case 449
			pcStrTitle = "No shipping"
			pcStrDetails = "<p>When the 'No Shipping' setting is active, the product is not considered a shipping product. The product weight, if any, is not added to the total weight for the order and if there are only 'No Shipping' products in the shopping cart, the shipping options page is skipped.</p><p>For example, a product available via electronic download or an hour of consulting service would be 'No Shipping' items.</p><p>The <u>no shipping text</u> displayed on the product details page can be edited by editing the file 'includes/languages.asp' (look for the string 'viewprd_8').</p><p>If you want to create a 'Free Shipping' promotion, you can use electronic coupons (discount codes) that make specific shipping options free. You can limit the applicability of those coupons to specific products, categories of products, order amounts, etc. Please see the User Guide for more details.</p>"	
			
		case 450
			pcStrTitle = "Tax by Zone"
			pcStrDetails = "<p>The Tax by Zone feature allows you to handle tax calculations correctly when complex scenarios apply. This feature was specifically created to assist <strong>Canadian</strong> users of ProductCart, but can certainly used elsewhere.</p><p>In Canada taxes on retail sales are calculated differently in different regions of the country. For example: <ul><li>Some provinces charge a Provincial Sales Tax (PST or RST for 'Retail Sales Tax'), but the rate is not the same in all the provinces.</li><li>All provinces also charge a federal tax on goods and services, called the 'GST'.</li><li>Nova Scotia, New Brunswick, and Newfoundland charge a Harmonized Sales Tax (HST) of 14 percent that combines both PST and GST.</li><li>In Quebec, the local tax is calculated on a total that includes the GST whereas in Ontario both PST and GST are calculated on the order total excluding other taxes.</ul></p><p>How can you correctly handle taxation when things get so complex? In ProductCart you can create zones (geographical locations that share the same tax) and specify a number of tax-related settings that apply to that zone. These settings are flexible enough for you to correctly calculate taxes in a place like Canada."	
			
		case 451
			pcStrTitle = "Using VAT (Value Added Tax)"
			pcStrDetails = "<p>Countries that use the Value Added Tax (VAT) can use this feature to show customers the tax that is included in the prices paid, and remove that tax when needed. VAT is shown on the product details page (to show the VAT incldued in the retail price) and during checkout (e.g. order verification page). The system behaves based on the following assumptions:</p><ul><li><strong>Prices are</strong> always entered <strong>VAT included</strong>.<br>When you add a product to your store, the product's price includes VAT. The system gives you the option to show the price without VAT on the product details page.</li><li>VAT rates are based on the country in which the store is located and on the product type.</li><li>Orders shipped to any address in the EU pay VAT.</li><li>Orders shipped to an address outside of the EU do not pay VAT, which is therefore removed from the order total. </li><li>The VAT rates that apply are those of the country in which the store is located (<a href='http://ec.europa.eu/taxation_customs/resources/documents/taxation/vat/how_vat_works/rates/vat_rates_en.pdf' target='_blank'>see current VAT rates</a>) </li><li>If a product is not assigned to a VAT Category the default VAT rate is used.</li></ul>"
			
		case 452
			pcStrTitle = "Using VAT Categories"
			pcStrDetails = "<p>ProductCart allows you to assign Products to different VAT Rate Categories based on the Product type. You can create as many categories as you need. For instance, you could assign all &quot;books&quot; to a Reduced Rate category, and all &quot;umbrellas&quot; to a Standard Rate category.&nbsp; Check the laws in your region if you're not sure which category a particular product belongs to. If all of your products fall into the same category, there is not need to use VAT Categories. Products that are not assigned to a VAT Category will automatically use the &ldquo;Default&rdquo; rate.&nbsp;</p>"
			
		case 453
			pcStrTitle = "Custom Search Fields: Order"
			pcStrDetails = "The order in which the field name is shown whenever a list of fields is displayed. Custom search fields are typically displayed in a drop-down menu."
			
		case 454
			pcStrTitle = "Custom Search Fields: Visibility on search pages"
			pcStrDetails = "This setting allows you to decide whether the search field should be shown on any advanced search page that exists in the Control Panel (CP Search) or in the storefront (SF Search).<br /><br />This option exists because there are cases in which you don't need to show a custom search field. Specifically, this is the case when you use custom search fields to add &quot;product properties&quot; to your products (e.g. MPN, UPC, ISBN, etc.) and use them as Export Fields to generate product data feeds for comparison shopping engines. <a href='http://wiki.earlyimpact.com/productcart/managing_search_fields#mapping_custom_search_fields_to_export_fields' target='_blank'>More information on this feature</a>."
			
		case 455
			pcStrTitle = "Custom Search Fields: Visibility on product details pages"
			pcStrDetails = "This setting allows you to decide whether the search field should be shown on the product details page in the storefront (SF Details) or on the Modify Product page in the Control Panel (CP Details).<br /><br />If the custom search field is used to store information that is not relevant to a customer in the storefront (e.g. UPC code), then you would want to choose not to show it in the storefront (SF Products option remains unchecked)."
			
		case 456
			pcStrTitle = "Locking a customer account"
			pcStrDetails = "The customer will be prevented from using his/her account. The customer will not be able to log in."
			
		case 457
			pcStrTitle = "Suspending a customer account"
			pcStrDetails = "The customer will be prevented from placing orders. The customer will be able to log in, but not place an order."
			
		case 460
			pcStrTitle = "Keyword Rich URLs"
			pcStrDetails = "This feature requires a server-side setting to be in place before it can be turned on.<br /><br /><strong>Do not turn on this feature</strong> until you have changed the 404 Error handler on your Web site. Please <a href='http://wiki.earlyimpact.com/productcart/seo-urls' target='_blank' onClick='JavaScript: window.close()'>review the documentation for details</a>."
			
		case 461
			pcStrTitle = "Keyword Rich URLs: path to 'Page Not Found' error page"
			pcStrDetails = "Enter the path to the custom page that you have created for &quot;Page Not Found&quot; error. That is, this is the page that will be loaded when somebody tries to load a page that does not exist, and that cannot be interpreted by the URL rewrite system used by the Keyword Rich URLs (or SEO URLs) feature.<br /><br />Example: <em>/404.html</em><br /><br />In this example an HTML file has been created and placed in the root of the Web site. Had the file been placed into a 'specialPages' folder, the field would contain: <em>/specialPages/404.html</em>.<br /><br />Note that this is just an example. The file could certainly have a different name and a different extension (e.g. <em>page_not_found.asp</em>)."
			
		case 462
			pcStrTitle = "Generate Navigation: table rows and cells vs. unordered lists"
			pcStrDetails = "<p>You can generate different HTML code for your category navigation: Web designers will likely prefer creating a category navigation that uses <strong>unordered lists </strong>(&lt;ul&gt;&lt;li&gt;My link&lt;/li&gt;&lt;/ul&gt;) rather than t<strong>able rows and cells </strong>(&lt;table&gt;&lt;tr&gt;&lt;td&gt;My link&lt;/td&gt;&lt;/tr&gt;&lt;/table&gt;), as they can more easily control unordered lists using cascading style sheets.<br><br>If you are not sure what to select, don't worry. Either setting will work. See which one you like best in your storefront.<br><br><strong>ProductCart v3</strong> by default created category navigation using table rows and cells. So if you upgraded from v3 and wish to keep using the same format, choose that option.</p>"	
			
		case 463
			pcStrTitle = "Shipping Surcharge"
			pcStrDetails = "<p>You can specify a surcharge to be added to the total shipping charges that will be calculated when this product is purchased. You can specify a different amount on the first unit (<strong>First Unit Surcharge</strong>), vs. additional units that they may purchase (<strong>Additional Unit(s) Surcharge</strong>). For more information on this feature, please <a href='http://wiki.earlyimpact.com/productcart/products_adding_new#shipping_surcharge_v4_only' target='_blank'>consult the ProductCart WIKI</a>.</p>"
			
		case 464
			pcStrTitle = "Store Settings: Disable 'Quick Buy' Button"
			pcStrDetails = "<p>ProductCart displays the 'Add to Cart' button wherever products are shown and available for purchase (<a href='http://www.earlyimpact.com/faqs/e-commerce_shopping_cart_questions.asp?faqid=466' target='_blank'>see when it is NOT shown</a>). You can disable this 'Quick Buy' feature and not show the 'Add to Cart' button in cases in which it would be shown.<br><br><u>NOTE</u>: On stores that are particularly busy or have a substantial amount of products, turning off this feature can also lead to a performance improvement as ProductCart will not have to determine whether the product is eligible for 'Quick Buy'.</p>"	
			
		case 465
			pcStrTitle = "Store Settings: Enable 'Stay on Page when Adding to Cart'"
			pcStrDetails = "<p>ProductCart automatically takes the buyer to the 'Show shopping cart' page (viewCart.asp) when one or more items are added to the cart. If you enable this feature, customers will <strong>stay on the same page</strong>. A message will be displayed to them indicating that the products were added to the cart. The message window contains a link to view the shopping cart.<br /><br />You can <strong>style the confirmation message window</strong> by using the styles contained in the cascading style sheet <em>pcStorefront.css</em>, which is located in the <em>pc</em> folder. Look for the styles toward the bottom of the page under &quot;Stay on Page when Adding to the Cart - Confirmation message window&quot;.</p>"	
			
		case 466
			pcStrTitle = "Store Settings: Checkout Options: Guest Checkout"
			pcStrDetails = "You have three options when it comes to allowing or not allowing customers to checkout as a 'Guest':<br /><br /><strong>Guest Checkout enabled</strong><br />Customers can opt to checkout without entering a password. The same customer can checkout on the same store, using the same e-mail address, multiple times, and multiple 'guest' profiles will be created. If the customer at any point decides to convert to a registered account, ProductCart will prompt him/her to consolidate into that account the other 'guest profiles' that exist in the database (accounts that use the same e-mail address).<br /><br /><strong>Guest Checkout allowed</strong><br />The difference between 'enabled' and 'allowed' is that with this second configuration the customer is prompted to enter a password during checkout. The customer can opt out and not do so. In other words, unlike with 'Guest Checkout Enabled', a password is initially required and the customer is prompted to enter it. So this may lead to more 'registered customers' than 'guests'. This behavior was introduced with ProductCart v4.1.<br /><br /><strong>Guest Checkout disabled</strong><br />A password is required. The system will not allow a customer to place an order unless a password has been entered. In this scenari, 'guests' cannot exist. Therefore, the system validates the e-mail address and prompt the customer to enter a different e-mail address or log into his/her account if the same address already exists in the database."
			
		case 467
			pcStrTitle = "Add/Edit Product: Image Magnifier"
			pcStrDetails = "<p>ProductCart uses a script called <a href=http://www.nihilogic.dk/labs/mojozoom/ target=_blank>MojoZoom</a> to provide image magnifying functionality to the product details page. This feature becomes useful only if you have a large product image and if customer can benefit from zooming into the image.</p>"	
			
		case 468
			pcStrTitle = "Store Settings: Restore saved shopping cart on next visit"
			pcStrDetails = "<p>When this feature is ON, a customer that visits a store that he/she has previously visited will experience the following behaviors:<ul><li>Products that had been added to the shopping cart will be added back to the cart</li><li>A pop-up message will notify the customer that those products were added back to the shopping cart (see example below)</li></ul>When the feature is OFF, both of these actions are not taken.</p><p align='center'><img src='images/pcv41_restored_cart.gif' alt='Restore saved cart' title='Example of modal window shown when a saved shopping cart is restored on the next visit' style='margin-top: 10px; margin-bottom: 10px;'></p><p>When a customer is logged in at the time products are added to the shopping cart, the information is saved to the store database together with the customer ID and the customer can access (and nickname) previously saved shopping carts through the customer account area. <a href='http://wiki.earlyimpact.com/productcart/savedshoppingcarts' target='_blank'>More information on Saved Shopping Carts</a>."	

		case 469
			pcStrTitle = "Incomplete Orders, Drop-offs, and Conversions"
			pcStrDetails = "<p><strong>Drop-off Rate</strong>: the number of <a href='resultsAdvancedAll.asp?B1=View+All&dd=1&OType=1' target='_blank'>incomplete orders</a> divided by the number of total orders. <br><em>It tells you how many customers decided not to buy after starting to checkout. <a href='http://wiki.earlyimpact.com/productcart/orders_status#incomplete_orders' target='_blank'>Learn more</a>.</em><p style='padding-top:10px;'><strong>Conversion Rate</strong>: the number of new customer registrations divided by the number of orders placed by new customers.<br><em>It tells you how many of your new customers ended up buying.</em></p><p style='padding-top:10px;'>The Drop-Off rate applies to all orders, whereas the Conversion Rate refers only to orders placed by new customers. Note that a &quot;new customer&quot; is not the same as a &quot;new visit&quot; to your Web site: only a subset of those who visit your Web store decide to begin the checkout process (and a new customer is saved into the store database).</p>"
			
		case 470
			pcStrTitle = "Store Settings: Google Analytics Profile ID"
			pcStrDetails = "<p>Google Analytics has been integrated into ProductCart starting with v4.5. Unlike with previous versions of ProductCart, you no longer need to add any code to the storefront. Simply enter the profile ID of the Web site that you wish to track, update the store settings, and you are done.</p><p><strong>To locate the profile ID</strong></p><ul><li>Log into Google Analytics</li><li>On the start page, locate the Web site that you wish to track: the Profile ID is shown right next to the URL and looks like this: UA-XXXXXX-X</li><li>It's unique for each Web site that you are tracking through Google Analytics (the last digit changes).</li></ul><p><strong>Removing existing tracking code</strong></p><p>If you are upgrading from a previous version of ProductCart, you might have previously added the Google Analytics code to your Web store footer (in either <em>footer.asp</em> or <em>inc_footer.asp</em>). Make sure to remove that code to avoid loading the tracking code twice. See the documentation for more details on this</p><p><a href='http://wiki.earlyimpact.com/widgets/integrations/googleanalytics/googleanalytics' target='_blank'>More information on the integration with Google Analytics</a>.</p>"	
			
		case 471
			pcStrTitle = "Store Settings: AddThis"
			pcStrDetails = "<p>AddThis provides buttons that allow visitors to your store to easily share a page on a social network, through Twitter, via e-mail, etc.</p><p>You can show the buttons either at the <strong>right of the page title</strong> (they are shown on the product details page, content pages, show specials, new arrivals, etc.) or <strong>below the Add To Cart button</strong>, on the product details page only.</p><p>For more information about AddThis and to select the buttons code that best fits your needs, <a href=http://www.addthis.com/web-button-select?where=website&clickbacks=1&type=bm&bm=tb5&analytics=1 target=_blank>click here</a>.</p><p><strong>NOTE: this feature utilizes a 3rd-party service. For technical support on issues related to the AddThis service, please contact AddThis support at <a href=http://www.addthis.com/forum/ target=_blank>http://www.addthis.com/forum/</a></strong></p>"	
			
		case 472
			pcStrTitle = "Store Settings: Company Logo"
			pcStrDetails = "<p>ProductCart uses this graphic in several printer-friendly pages (e.g. store invoices). By default, these pages show a 'Your Logo Here' image. Here you can specify another image. In the 'Company logo' field, enter the file name (no file path) or click on the search icon to locate the image. ProductCart assumes that the file is in the 'pc/catalog' directory on your Web server.</p><p>If you have not yet uploaded the image to your store, you can do so now by clicking on the upload icon.</p>"
			
		case 473
			pcStrTitle = "Store Settings: Default Meta Tags"
			pcStrDetails = "You can enter default Meta Tags that will be used on pages for which page-specific meta tags (categories, products, content pages) are not generated dynamically by ProductCart.</p><p>Please note that both the Keywords and the Description meta tags are not used by Google to determine a Web site's ranking. The Description meta tag, however, is often used to describe the Web site in the search results, and therefore it remains an important page element.</p><p>If you are interested in learning more about Meta Tags, you will find a lot of articles on the Internet. For example, see the <a href='http://blog.earlyimpact.com/2011/01/google-seo-starter-guide-and.html' target='_blank'>Google SEO Starter Guide &gt;&gt;</a>."
			
		case 474
			pcStrTitle = "Shipping Settings: Maximum Weight Per Package"
			pcStrDetails = "<p>Products that are set as &quot;oversized&quot; are always treated as separate packages when calculating shipping charges. For products that are not set as oversized, ProductCart can <u>automatically divide the shipment over multiple packages</u>.</p><p>The number of packages that will be shipped is calculated according to the following formula:</p><p><blockquote><strong>Total order weight / Max weight per package = Number of packages to be shipped</strong></blockquote></p>"

		case 475
			pcStrTitle = "Shipping Settings: Residential vs. Commercial Address"
			pcStrDetails = "<p>Shipping providers such as UPS and FedEx charge higher rates for Residential deliveries compared to Commercial deliveries. You can choose to let customers specify which type of address they are shipping to (<em>this was the only behavior supported before ProductCart v4.5</em>) or automatically make a selection for them.</p><p>You can also set the system to automatically default to a residential address when the customer is a new or retail customer, and to a commercial address when the customer is logged in as a wholesale customer."

		case 476
			pcStrTitle = "E-mail settings: Use 'From' E-mail in 'Contact Us' form"
			pcStrDetails = "The 'Contact Us' page in the storefront (contact.asp) by default uses the customer's e-mail as the 'From' e-mail address. That way, if you reply to the message, the message will be sent to the customer.<br><br>Check this box if your hosting provider blocks e-mails sent from a different domain: the 'Contact Us' form will use the store's e-mail instead.<br><br>Some hosting companies block any message that is sent through the Web site if the Sender's domain is different than the web site's domain, for added security.<br><br>If you check with checkbox, you will need to manually copy the customer's e-mail address into the Recipient field when you reply to a message coming in from the Contact Form, otherwise you will be sending a message back to yourself!" 
			
		case 477
			pcStrTitle = "Percent Amount is applied against original Order Amount"
			pcStrDetails = "When you enter a value and Adjust the Commission Earned based on a Percentage... the percentage value is applied against the original Order Amount! E.G. if the Order Amount is $300 and you enter a value of 50 and choose 'Percent(%), the Commission Earned will be $150. If you enter a value of 50 and choose 'Amount($), the Commission Earned will be $50." 

			
		' Build To Order
		
		case 500
			pcStrTitle = "Build To Order: &quot;Default&quot; and &quot;Base&quot; prices"
			pcStrDetails = "A Build To Order product's &quot;<strong>Default Price</strong>&quot; is the price shown when customers browse the store (e.g. the price shown in a search, or on the Specials or New Arrivals pages).<br /><br />This price is calculated as <strong>Base Price</strong> + <strong>Default Option Prices</strong>, where the last is the sum of the prices of the items selected as &quot;default&quot; when you assign configurable items to the BTO product (Locate a BTO Product &gt; Configure).<br /><br />The Base Price is set when adding/modifying a BTO product. It can also be set or edited for multiple BTO products at once using the &quot;<a href=""updBTOPrdPrices.asp"" target=""_blank"">Update Base Prices</a>&quot; feature." 


		
		' Apparel Add-on
		
		case 600
			pcStrTitle = "Apparel Add-on: sub-product prices and pricing categories"
			pcStrDetails = "If you are using pricing categories, you can set prices at the sub-product level for each of the pricing categories used on your store. Please note the following about the 'Manage Sub-Product Prices' page: <ul><li style='padding-bottom: 5px;'>Prices = total price for that sub-products.<br>Unlike on the page where you Manage Sub-Products, this is <u>not</u> a price differential applied to the retail or wholesale price. This is the actual price paid for that sub-product by a customer belonging to that pricing category.</li><li style='padding-bottom: 5px;'>The default price is shown in blue when the pricing category is &quot;Across The Board&quot;. Leave the value in the input field as &quot;0&quot; (or change it to &quot;0&quot;) if you want to use the default price</li><li style='padding-bottom: 5px;'>The default price is not shown when a category is &quot;Product by Product&quot;</li></ul>" 
			
			
		' EIG
		
		case 700
			pcStrTitle = "NetSource Commerce Gateway"
			pcStrDetails = "<p><strong>REFUND</strong><br /><br />When you use the NetSource Commerce Payment Gateway, you can refund a payment directly from the ProductCart Control Panel. A &quot;Refund&quot; button will be available in the &quot;Payment Status&quot; tab of the &quot;Order Details&quot; page when the payment status is &quot;Paid&quot;.<br><br>When you refund a payment, the payment gateway is contacted and the payment is refunded. In ProductCart, the payment status becomes &quot;Refunded&quot;.<br><br>The order processing status remains &quot;as is&quot; because ProductCart does not make an assumption on the circumstances of the refund (e.g. you may or may not need to cancel the order).</p>" 
			
			
		'PayPal Advanced
		
		case 800
			pcStrTitle = "PayPal Payments Advanced"
			pcStrDetails = "<p><strong>Partner Name:</strong> Your partner name is PayPal.</p><p><strong>Merchant Login:</strong> This is the login name you created when signing up for PayPal Payments Advanced.</p><p><strong>User:</strong> PayPal recommends entering a User Login here instead of your Merchant Login. You can set up a User profile in <a href='https://manager.paypal.com'>PayPal Manager</a>. This will enhance security and prevent service interruption should you change your Merchant Login password.</p><p><strong>Password:</strong> This is the password you created when signing up for PayPal Payments Advanced or the password you created for API calls.</p><p><strong>Test Mode: </strong>Visit <a href='https://developer.paypal.com' target='_blank'>PayPal's Developer Site</a> to obtain &quot;SandBox&quot; (Test) credentials. Do not use your real account information while test mode is enabled.</p>"

		case 801
			pcStrTitle = "PayPal Payflow Link"
			pcStrDetails = "<p><strong>Partner Name:</strong> This should be the same partner name that's used when logging into your PayPal Payflow account.</p><p><strong>Merchant Login:</strong> This is the login name you created when signing up for Payflow.</p><p><strong>User:</strong> PayPal recommends entering a User Login here instead of your Merchant Login. You can set up a User profile in <a href='https://manager.paypal.com'>PayPal Manager</a>. This will enhance security and prevent service interruption should you change your Merchant Login password.</p><p><strong>Password:</strong> This is the password you created when signing up for Payflow or the password you created for API calls.</p><p><strong>Test Mode: </strong>Visit <a href='https://developer.paypal.com' target='_blank'>PayPal's Developer Site</a> to obtain &quot;SandBox&quot; (Test) credentials. Do not use your real account information while test mode is enabled.</p>"
		
		case 802
			pcStrTitle = "PayPal Payments Standard"
			pcStrDetails = "<p><strong>PayPal Account ID/Email Address:</strong> Enter the email address associated with your PayPal account.</p><p><strong>Test Mode: </strong>Visit <a href='https://developer.paypal.com' target='_blank'>PayPal's Developer Site</a> to obtain &quot;SandBox&quot; (Test) credentials. Do not use your real account information while test mode is enabled.</p>"
			
		case 803
			pcStrTitle = "PayPay Express"
			pcStrDetails = "<p><strong>Test Mode: </strong>Visit <a href='https://developer.paypal.com' target='_blank'>PayPal's Developer Site</a> to obtain &quot;SandBox&quot; (Test) credentials. Do not use your real account information while test mode is enabled.</p>"
			
		case 804
			pcStrTitle = "Payflow Pro"
			pcStrDetails = "<p><strong>Partner Name:</strong> This should be the same partner name that's used when logging into your PayPal Payflow account.</p><p><strong>Merchant Login:</strong> This is the login name you created when signing up for Payflow.</p><p><strong>User:</strong> PayPal recommends entering a User Login here instead of your Merchant Login. You can set up a User profile in <a href='https://manager.paypal.com'>PayPal Manager</a>. This will enhance security and prevent service interruption should you change your Merchant Login password.</p><p><strong>Password:</strong> This is the password you created when signing up for Payflow or the password you created for API calls.</p><p><strong>Test Mode: </strong>Visit <a href='https://developer.paypal.com' target='_blank'>PayPal's Developer Site</a> to obtain &quot;SandBox&quot; (Test) credentials. Do not use your real account information while test mode is enabled.</p>"
			
		case 805
			pcStrTitle = "PayPay Payments Pro"
			pcStrDetails = "<p><strong>Test Mode: </strong>Visit <a href='https://developer.paypal.com' target='_blank'>PayPal's Developer Site</a> to obtain &quot;SandBox&quot; (Test) credentials. Do not use your real account information while test mode is enabled.</p>"
			
		case 806
			pcStrTitle = "Website Payments Pro"
			pcStrDetails = "<p><strong>Partner Name:</strong> This should be the same partner name that's used when logging into your PayPal Payflow account.</p><p><strong>Merchant Login:</strong> This is the login name you created when signing up for Payflow.</p><p><strong>User:</strong> PayPal recommends entering a User Login here instead of your Merchant Login. You can set up a User profile in <a href='https://manager.paypal.com'>PayPal Manager</a>. This will enhance security and prevent service interruption should you change your Merchant Login password.</p><p><strong>Password:</strong> This is the password you created when signing up for Payflow or the password you created for API calls.</p><p><strong>Test Mode: </strong>Visit <a href='https://developer.paypal.com' target='_blank'>PayPal's Developer Site</a> to obtain &quot;SandBox&quot; (Test) credentials. Do not use your real account information while test mode is enabled.</p>"
			
	end select
%>