
<%


  nome=request.querystring("Nome")
  email = request.querystring("CFN")


'########################
'## SET DATABASE PATH ###
'########################

   DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../database/petition.mdb")



   SQLupd="UPDATE firmatari SET conferma=1 where email='" & email & "'"

   dim rstemp, conntemp
   set conntemp=server.createobject("adodb.connection")
   conntemp.open DSN
   set rstemp=conntemp.execute(SQLupd)





'##############################################################
'## LETTER CONFIRMING THE CORRECT RECORDING OF THE SIGNATURE ##
'##############################################################

  lettera="Dear " & nome & vbCrLf & vbCrLf
  lettera=lettera & "Thanks for signing our petition" & vbCrLf
  lettera=lettera & "Your signature is now correctly recorded!" & vbCrLf & vbCrLf
  lettera=lettera & "Thank you!" & vbCrLf & vbCrLf
  lettera=lettera & "Webmaster" & vbCrLf
  lettera=lettera & "http://www.retrogas.com" & vbCrLf & vbCrLf & vbCrLf & vbCrLf


'########################################################
'## SET YOUR SMTP SERVER AND PETITION'S E-MAIL ADDRESS ##
'########################################################

Dim ObjMail
Set ObjMail = Server.CreateObject("CDONTS.NewMail")
ObjMail.To = email
'in the next line you can change that email address to something else
'like "myform@myurl.com", just be sure and put it between quotes " "
ObjMail.From = "Webmaster@retrogas.com"
'you can also change "Form Submission" to something else like "form results" etc.
ObjMail.Subject = "Please confirm your signature!"
ObjMail.Body = lettera
ObjMail.Send
Set ObjMail = Nothing

%>
<html>

<head>
<title>www.retrogas.com</title>
</head>

<body bgcolor="#FFFFFF">
<div align="center"><center>

<table border="0" width="468" cellspacing="0" cellpadding="0"
style="border: 1px dashed rgb(0,0,0)">
  <tr>
    <td width="100%"><p align="center"><img src="images/petition.gif" width="468" height="60"></td>
  </tr>
  <tr>
    <td width="100%"><p align="center">&nbsp;</p>
    <p align="center">&nbsp;</p>
    <p align="center"><strong><font face="Verdana" color="#008000"><small><small>Your
    signature has been recorded!</small></small></font></strong></p>
    <p align="center">&nbsp;</p>
    <p align="center"><a href="default.asp" style="color: rgb(0,0,255)" target="_top"><font
    face="Verdana"><small><small>Homepage</small></small></font></a></p>
    <p align="center"><small><small><font face="Verdana">&nbsp;</font></small></small></p>
    <p align="center"><font face="Verdana"><br>
    <font color="#FFFFFF"></font></font></td>
  </tr>
</table>
</center></div>
</body>
</html>
