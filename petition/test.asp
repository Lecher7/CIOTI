<form method=Post action=get.asp>
FName<br>
<input type="text" name="t1" size="20"><br>
LName<br>
<input type="text" name="t2" size="20"><br>
Email<br>
<input type="text" name="t3" size="20"><br>
City<br>
<input type="text" name="t4" size="20"><p>
<input type=submit value=Submit></form> 


Copy this into notepad and save as "get.asp" <%
Dim t1name,t1,t2name,t2,t3name,t3,t4name,t4
t1name = "FName"
t1 = Request.Form("t1")
t2name = "LName"
t2 = Request.Form("t2")
t3name = "Email"
t3 = Request.Form("t3")
t4name = "City"
t4 = Request.Form("t4")
Dim ObjMail
Set ObjMail = Server.CreateObject("CDONTS.NewMail")
ObjMail.To = "stevedetmer@yahoo.com"
'in the next line you can change that email address to something else 
'like "myform@myurl.com", just be sure and put it between quotes " "
ObjMail.From = "stevedetmer@yahoo.com"
'you can also change "Form Submission" to something else like "form results" etc.
ObjMail.Subject = "Form Submission"
ObjMail.Body = t1name & vbcrlf&_
t1 & vbcrlf&_
t2name & vbcrlf&_
t2 & vbcrlf&_
t3name & vbcrlf&_
t3 & vbcrlf&_
t4name & vbcrlf&_
t4
ObjMail.Send
Set ObjMail = Nothing
'HERE you make a choice. You can redirect the user to any page in your site
Response.Redirect "http://pacosdrivers.com"
'Or just say thanks. Delete the line you dont want. Either above or below
Response.Write"Thank You"
%> 

