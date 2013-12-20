						<% 
						' This file is included in most payment gateway files
						' It provides the code for the "Back" and "Place Order" buttons shown at the
						' bottom of the payment form.
						
							If scSSL="1" And scIntSSLPage="1" Then
								tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/OnePageCheckout.asp?PayPanel=1"),"//","/")
								tempURL=replace(tempURL,"https:/","https://")
								tempURL=replace(tempURL,"http:/","http://")
							Else
								tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/OnePageCheckout.asp?PayPanel=1"),"//","/")
								tempURL=replace(tempURL,"https:/","https://")
								tempURL=replace(tempURL,"http:/","http://")
							End If
						%> 
						<a href="<%=tempURL%>"><img src="<%=rslayout("back")%>"></a>&nbsp;&nbsp; 
						<input type="image" name="Continue" src="<%=rslayout("pcLO_placeOrder")%>" id="submit">
						<script type="text/javascript">
                            $(document).ready(function() {
                                $('#submit', this).attr('disabled', false);
                                $('form').submit(function(){
                                    $('#submit', this).attr('disabled', true);
                                    return 
                                });
                            });
                        </script> 