<%
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma", "No-Cache"

Dim Conexao, SQL, Tabela, Str, Exec, Mat, MSG, User, usuario, anotacao

'User = Trim(Session("Matricula"))
User = Session("usuario")
'Mat = Trim(request.querystring("Mat"))
'Mat = Trim(Request("txtMatricula"))
Exec = Trim(request.querystring("Exec"))
MSG = Trim(request.querystring("MSG"))
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Altera Senha</title>
<link rel="stylesheet" href="include/pgo.css" type="text/css">
<script language="javascript" src="include/ValidaNumero.js"></script>
<script language=javascript>
<!--
	function ValidaDados() {
		if (frmDados.txtAnotacao.value == "") {
			alert("Preencha o campo Anotação.");
			frmDados.txtAnotacao.focus();
			return false;
		}
				frmDados.action = "anotacao.asp?User="+frmDados.txtUser.value+"&Exec=S";
				frmDados.submit();
		}
-->
</SCRIPT>
</head>
<body onLoad="javascript:document.frmDados.txtNova.focus();">
<hr color="#002b4c" size="1">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
  <tr>
    <td width="100%" colspan="2">
      <p align="center"><font face="Verdana" size="3"><b>Bloco de Notas</b></font></td>
  </tr>
</table>
<form method="POST" name="frmDados">
<%
Sub Localizar
%>
<!--#include file="AbreConexao.asp"-->
<%
SQLL = "SELECT * FROM Vendas_ProspectOi..SIS_Anotacao WHERE usuario = '" & User & "'"
Set RSBUSCA = server.createobject("ADODB.Recordset")
'response.write SQLL
RSBUSCA.Open SQLL, Conexao

If Not RSBUSCA.EOF Then

	usuario = RSBUSCA("usuario")
	anotacao = RSBUSCA("Anotacao")

End If
Conexao.Close
Set SQL = Nothing
Set RSBUSCA = Nothing
End Sub

If ID <> "" and Enviar <> "S" Then
	Localizar
End If
%>
<table border="0" width="100%">
  <tr>
    <td width="5%" valign="top"><font face="Verdana" size="1"><b>Anotação:</b></font></td>
    <td width="95%"><textarea name="txtAnotacao" cols="100" rows="20" class="formfield"><%=anotacao%></textarea></td>
  </tr>
  <tr>
    <td width="5%"><b><font face="Verdana" size="1">&nbsp;</font></b></td>
    <td width="95%"></td>
  </tr>
      <input type="hidden" name="txtUser" value="<%=User%>">
  <tr>
    <td colspan="2">
    <p><input name="Salvar" type="submit" onClick="return ValidaDados();" id="Salvar" value="Salvar" class="formbutton">&nbsp;<input class="formbutton" type="reset" value="Limpar" name="B2"></p>
    </td>
  </tr>
</table>
</form>
<%
Sub Altera
%>
<!--#include file="AbreConexao.asp"-->
<%
SQL = "SELECT * FROM SIS_Usuarios WHERE Usuario = '" & User & "'"
Set Tabela = Conexao.execute(SQL)

If Tabela.EOF and Tabela.BOF Then
	Response.Write "<script language='javascript'>"
	Response.Write "alert('Anotação Salva com Sucesso!');"
	Response.Write "parent.parent.top.location.href='anotacao.asp';"
	Response.Write "</script>"
	Response.End
Else
	SQL = "UPDATE Vendas_ProspecOi..SIS_Anotacao SET usuario = '" & Trim(Request("txtNova")) & "', anotacao = '" & Trim(Request("txtAnotacao")) & "' WHERE usuario = '" & User & "'"
	Set Tabela = Conexao.execute(SQL)
	If Err.number = 0 Then
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Anotação SAlva com Sucesso');"
		Response.Write "parent.parent.top.location.href='alteraOk.asp';"
		Response.Write "</script>"
		Response.End
	Else
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Ocorreu um erro ao modificar sua senha. Favor contatar o Administrador');"
		Response.Write "</script>"
		Response.End
	End If
End If

	Tabela.Close
	Set SQL = Nothing
	Set Conexao = Nothing
End Sub

If Exec = "S" Then
	Altera
End If
%>
</body>

</html>
