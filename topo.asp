<%
Dim User, Usuario, Matric, Acesso
User = Trim(Session("usuario"))

If User = "" Then
	Response.Write "<script language='javascript'>"
	Response.Write "alert('Efetue o logon novamente.');"
	Response.Write "parent.parent.top.location.href='default.asp';"
	Response.Write "</script>"
	Response.End
End If

Sub Busca_Usuario
%>
<!--#include file="include/AbreConexao.asp"-->
<%
SQL = " SELECT Nome, usuario, acesso FROM SIS_Usuarios "
SQL = SQL & " WHERE usuario = '" & Trim(Session("usuario")) & "' "
Set Tabela = Conexao.execute(SQL)

If Tabela.EOF and Tabela.BOF Then
	Response.Write "<script language='javascript'>"
	Response.Write "alert('Este usuário não está cadastrado. Favor entrar em contato com o Administrador.');"
	Response.Write "parent.parent.top.location.href='sair.asp';"
	Response.Write "</script>"
	Response.End
Else
	Usuario = Tabela.Fields("Nome").value
	Matric = Tabela.Fields("usuario").value
	Acesso = Tabela.Fields("acesso").value
End If

Tabela.Close
Set SQL = Nothing
Set Conexao = Nothing
End Sub

If User <> "" Then
	Busca_Usuario
End If
%>
<html>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>MudEnd</title>
<base target="principal">
<link rel="stylesheet" href="include/pgo.css" type="text/css">
</head>

<body text="#000000" leftmargin="5" topmargin="5" marginwidth="0" marginheight="0">
<table border="0" width="100%" cellspacing="1">
  <tr>
   <td align="center" width="5%"><b>Usuário: <font color="blue"><%=Usuario%></font></b>
     <td width="10%" height="38"><b>&nbsp;<img border="0" src="imagens/img_sist_seta.gif"><a class="topo" target="_top" href="entrada.asp" onMouseOver="window.status='Menu Principal';return true;" title="Voltar para página principal">
      Principal</a></b></td>
    <td width="5%">
      <b><img border="0" src="imagens/img_sist_seta.gif"><a class="topo" target="_top" href="sair.asp" onMouseOver="window.status='';return true;" title="Sair">
        Sair</a></b>
    </td>
    <% If Acesso <> "Vendas" Then %>
   <% Else %>

    <% End If %>
  </tr>
</table>
</body>

</html>