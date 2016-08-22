<%
Dim User, MSG, pagina
User = Trim(Session("usuario"))
MSG = Trim(request.querystring("MSG"))

If User = "" Then
	Response.Write "<script language='javascript'>"
	Response.Write "alert('Efetue o logon novamente.');"
	Response.Write "parent.parent.top.location.href='default.asp';"
	Response.Write "</script>"
	Response.End
End If

If Trim(MSG) = "Sim" Then
	pagina = "altera.asp?MSG=Sim"
ElseIf User = "59371" then
	pagina = "menu/menu.asp"
Else
	pagina = "menu/menu.asp"
End If
%>
<html>
<head>
<title>DNA - <%=Session("Nome")%> - <%=Session("acesso")%></title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<frameset framespacing="0" border="0" rows="40,*" frameborder="0">
  <frame name="cabecalho" scrolling="no" noresize target="principal" src="topo.asp?User=<%=User%>">
  <frame name="principal" scrolling="auto" src="<%=pagina%>">
  <noframes>
  <body topmargin="0" leftmargin="0">

  <p>Esta página usa quadros mas seu navegador não aceita quadros.</p>

  </body>
  </noframes>
</frameset>

</html>
