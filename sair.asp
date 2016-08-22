<%
Session("usuario") = ""

	Response.Write "<script language='javascript'>"
	Response.Write "alert('Logoff efetuado com sucesso!!');"
	Response.Write "parent.parent.top.location.href='default.asp';"
	Response.Write "</script>"
	Response.End
%>