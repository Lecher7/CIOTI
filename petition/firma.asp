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
    <td width="100%"><p align="center"><small><font face="Verdana" color="#FFFFFF"><small><small>.</small></small></font></small><font
    face="Verdana" color="#000000"><strong><br>
    <small><small>Sign the petition!</small></small></strong></font></p>
    <p><script Language="JavaScript">
<!--
function Validator(theForm)
{

  if (theForm.NOME.value == "")
  {
    alert("Please insert your name.");
    theForm.NOME.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþ' \t\r\n\f";
  var checkStr = theForm.NOME.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Not admitted caracters in the NAME field.");
    theForm.NOME.focus();
    return (false);
  }

  if (theForm.COGNOME.value == "")
  {
    alert("Please insert your surname.");
    theForm.COGNOME.focus();
    return (false);
  }


  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþ' \t\r\n\f";
  var checkStr = theForm.COGNOME.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Not admitted caracters in the SURNAME field.");
    theForm.COGNOME.focus();
    return (false);
  }

  if (theForm.EMAIL.value == "")
  {
    alert("Not valid e-mail address.");
    theForm.EMAIL.focus();
    return (false);
  }


if (theForm.CFNEMAIL.value != theForm.EMAIL.value)
  {
    alert("E-mail doesn't match.");
    theForm.EMAIL.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþ0123456789-_-.@$£&+";
  var checkStr = theForm.EMAIL.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Not admitted caracters in the E-MAIL field.");
    theForm.EMAIL.focus();
    return (false);
  }

  if (theForm.CITTA.value == "")
  {
    alert("Please insert you town.");
    theForm.CITTA.focus();
    return (false);
  }


  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþ0123456789-' \t\r\n\f";
  var checkStr = theForm.CITTA.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Not admitted caracters in the TOWN field.");
    theForm.CITTA.focus();
    return (false);
  }

  if (theForm.PROFESSIONE.value == "")
  {
    alert("Please insert your job.");
    theForm.PROFESSIONE.focus();
    return (false);
  }


  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþ0123456789-' \t\r\n\f";
  var checkStr = theForm.PROFESSIONE.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Not admitted caracters in the JOB field.");
    theForm.PROFESSIONE.focus();
    return (false);
  }

  if (theForm.NAZIONE.value == "")
  {
    alert("Please insert your country.");
    theForm.NAZIONE.focus();
    return (false);
  }


  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþ0123456789- \t\r\n\f";
  var checkStr = theForm.NAZIONE.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Not admitted caracters in the COUNTRY field.");
    theForm.NAZIONE.focus();
    return (false);
  }
  return (true);
}
//-->
</script></p>
    <form method="POST" action="caricadb.asp" onsubmit="return Validator(this)">
      <div align="center"><center><table border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td align="right" bgcolor="#FFFFFF"><small><small><font face="Verdana" color="#000000"><strong>First Name:</strong></font><font
          face="Verdana" color="#FF0000">*</font></small></small></td>
          <td bgcolor="#FFFFFF"><font face="Verdana"><input type="text" name="NOME" size="20"
          maxlength="50" style="font-family: Verdana; font-size: 10"></font></td>
        </tr>
        <tr>
          <td align="right" bgcolor="#FFFFFF"><small><small><font face="Verdana" color="#000000"><strong>Last Name:</strong></font><font
          face="Verdana" color="#FF0000">*</font></small></small></td>
          <td bgcolor="#FFFFFF"><font face="Verdana"><input type="text" name="COGNOME" size="20"
          maxlength="50" style="font-family: Verdana; font-size: 10"></font></td>
        </tr>
        <tr>
          <td align="right" bgcolor="#FFFFFF"><small><small><font face="Verdana" color="#000000"><strong>E-mail:</strong></font><font
          face="Verdana" color="#FF0000">*</font></small></small></td>
          <td bgcolor="#FFFFFF"><font face="Verdana"><input type="text" name="EMAIL" size="20"
          maxlength="200" style="font-family: Verdana; font-size: 10"></font></td>
        </tr>
        <tr>
          <td align="right" bgcolor="#FFFFFF"><small><font face="Verdana" color="#000000"><strong><small>Confirm</small><br>
          <small>E-mail:</strong></font><font face="Verdana" color="#FF0000">*</font></small></small></td>
          <td bgcolor="#FFFFFF"><font face="Verdana"><input type="text" name="CFNEMAIL" size="20"
          maxlength="200" style="font-family: Verdana; font-size: 10"></font></td>
        </tr>
        <tr>
          <td align="right" bgcolor="#FFFFFF"><small><small><font face="Verdana" color="#000000"><strong>City:</strong></font><font
          face="Verdana" color="#FF0000">*</font></small></small></td>
          <td bgcolor="#FFFFFF"><font face="Verdana"><input type="text" name="CITTA" size="20"
          maxlength="50" style="font-family: Verdana; font-size: 10"></font></td>
        </tr>
        <tr>
          <td align="right" bgcolor="#FFFFFF"><small><small><font face="Verdana" color="#000000"><strong>State:</strong></font><font
          face="Verdana" color="#FF0000">*</font></small></small></td>
          <td bgcolor="#FFFFFF"><font face="Verdana"><input type="text" name="PROFESSIONE" size="20"
          maxlength="50" style="font-family: Verdana; font-size: 10"></font></td>
        </tr>
        <tr>
          <td align="right" bgcolor="#FFFFFF"><small><small><font face="Verdana" color="#000000"><strong>Country:</strong></font><font
          face="Verdana" color="#FF0000">*</font></small></small></td>
          <td bgcolor="#FFFFFF"><font face="Verdana"><input type="text" name="NAZIONE" size="20"
          maxlength="50" style="font-family: Verdana; font-size: 10"></font></td>
        </tr>
        <tr>
          <td colspan="2" bgcolor="#FFFFFF"><div align="center"><center><p><font face="Verdana"
          color="#FF0000"><small><small>* mandatory field</small></small></font></td>
        </tr>
        <tr align="center">
          <td colspan="2" bgcolor="#FFFFFF"><font face="Verdana"><div align="center"><center><p><input
          type="submit" style="font-weight: bold; font-family: Verdana; font-size: 10" value="Sign!"></font></td>
        </tr>
      </table>
      </center></div>
    </form>
    <p align="center"><font face="Verdana" color="#FF0000"></font><br>
    <font face="Verdana"><br>
    <font color="#FFFFFF"></font></font></td>
  </tr>
</table>
</center></div>
</body>
</html>
