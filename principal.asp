<%
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma", "No-Cache"


User = Session("usuario")
If User = "" Then
	Response.Write "<script language='javascript'>"
	Response.Write "alert('Efetue o logon novamente.');"
	Response.Write "parent.parent.top.location.href='default.asp';"
	Response.Write "</script>"
	Response.End
End If

%>
<html>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>BS Digital</title>
<link rel="stylesheet" href="../include/pgo.css" type="text/css">
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="initialize();calcRoute()">
<table width="100%" border="0" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr>
    <td width="25%" align="left" valign="top">
    <IFRAME src="menu/menu.asp" width="450" height="1150" scrolling="no" frameborder="0" align="left"></IFRAME> 
 </td>
    <td width="30%" align="center" valign="top">
    <IFRAME src="mensagem.asp" width="300" height="300" scrolling="no" frameborder="0" align="center"></IFRAME> 
    </td>

    <td width="30%" align="center" valign="top">
 <IFRAME src="bs_Atualizacao.asp" width="300" height="300" scrolling="no" frameborder="0" align="center"></IFRAME>	
	</td>
  </tr>

</table>
</body>

