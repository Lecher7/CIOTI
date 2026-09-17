<%
'****************************************************************************************
'**  Copyright Notice
'**  Copyright 2003 DUware All Rights Reserved.
'**  This program is free software; you can modify (at your own risk) any part of it
'**  under the terms of the License that accompanies this software and use it both
'**  privately and commercially.
'**  All copyright notices must remain in tacked in the scripts and the
'**  outputted HTML.
'**  You may use parts of this program in your own private work, but you may NOT
'**  redistribute, repackage, or sell the whole or any part of this program even
'**  if it is modified or reverse engineered in whole or in part without express
'**  permission from the author.
'**  You may not pass the whole or any part of this application off as your own work.
'**  All links to DUware and powered by logo's must remain unchanged and in place
'**  and must remain visible when the pages are viewed unless permission is first granted
'**  by the copyright holder.
'**  This program is distributed in the hope that it will be useful,
'**  but WITHOUT ANY WARRANTY; without even the implied warranty of
'**  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE OR ANY OTHER
'**  WARRANTIES WHETHER EXPRESSED OR IMPLIED.
'**  No official support is available for this program but you may post support questions at: -
'**  http://www.duware.com/support
'****************************************************************************************
%>
<%'-------------------------------------------------------BEGIN OF CONFIGURATION----------------------------------------------%>

<%
'YOUR PAYPAL ACCOUNT ID (EMAIL ADDRESS)
myPaypalID = "ozzyrux@softhome.net"


'YOUR PAYPAL CURRENCY SETTING (EURO, CAN, USD, POUND AND YEN ONLY)
myPaypalCurrency = "USD"
myPaypalCurrencySign = "$"


'THE URL THAT BUYERS WILL BE REDIRECTED AFTER PAYMENT IS MADE AT PAYPAL
myReturnURL = "http://www.northeastcards.com"

'THE URL THAT BUYERS WILL BE REDIRECTED IF CANCEL THE PAYMENT PROCESS AT PAYPAL
myCancelURL = "http://www.northeastcards.com"
%>

<%'-------------------------------------------------------END OF CONFIGURATION----------------------------------------------%>


















<%
Function DoSpace(str)
  DoSpace = (Replace(str, vbCrlf, "<br>"))
End Function
%>