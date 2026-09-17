
<%

  nome=replace(request.form("NOME"),"'","''")
  cognome=replace(request.form("COGNOME"),"'","''")
  email = request.form("EMAIL")
   citta = replace(request.form("CITTA"),"'","''")
	professione = replace(request.form("PROFESSIONE"),"'","''")
    nazione = replace(request.form("NAZIONE"),"'","''")


'########################
'## SET DATABASE PATH ###
'########################

   DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../database/petition.mdb")

   SQLupd="INSERT INTO firmatari (Nome,Cognome,Email,Citta,Professione,Nazionalita) VALUES ('" & nome & "','" & cognome & "','" & email & "','" & citta & "','" & professione & "','" & nazione & "')"

   dim rstemp, conntemp
   set conntemp=server.createobject("adodb.connection")
   conntemp.open DSN
   set rstemp=conntemp.execute(SQLupd)


'#################################################
'## LETTER ASKING TO CONIRM SIGNATURE WITH LINK ##
'#################################################

  lettera="Dear " & nome & vbCrLf & vbCrLf
  lettera=lettera & "Thanks for signing the petition" & vbCrLf & vbCrLf
  lettera=lettera & "Please confirm your signature by clicking the following link: http://www.progtunes.com/retrogas/petition/conferma.asp?cfn=" & email & "&Nome=" & nome & vbCrLf & vbCrLf & vbCrLf
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
<title>www.RetroGas.com</title>
</head>

<body bgcolor="#FFFFFF">
<div align="center"><center>

<table border="0" width="468" cellspacing="0" cellpadding="0"
style="border: 1px dashed rgb(0,0,0)">
  <tr>
    <td width="100%"><p align="center"><img src="images/petition.gif" width="468" height="60"></td>
  </tr>
  <tr>
    <td width="100%"><p align="center"><small><font face="Verdana" color="#FFFFFF"><small><small>.</small></small></font></small><strong><font
    face="Verdana" color="#000000"><br>
    </font><font face="Verdana" color="#FF0000"><small><small>Your signature is waiting to be
    confirmed!</small></small></font></strong></p>
    <p align="center"><font face="Verdana"><small><small>You will soon get an e-mail
    containing a special link you have</small></small><br>
    <small><small>to click to finally confirm your signature.</small></small></font></p>
    <p align="center"><font face="Verdana"><small><small>Thank you!</small></small></font></p>
    <p align="center"><font face="Verdana"><small><a href="default.asp"
    style="color: rgb(0,0,255)"><small>Homepage</small></a></small></font></p>
    <p align="center"><small><small><font face="Verdana">&nbsp;</font></small></small></p>
    <p align="center"><font face="Verdana"><small><small>Powered by <a
    href="http://www.chimicon.com/petition" target="_blank" style="color: rgb(0,0,0)">Petition
    Script 2.0</a></small></small><br>
    <font color="#FFFFFF"><small><small>.</small></small></font></font></td>
  </tr>
</table>
</center></div>
</body>
</html>
