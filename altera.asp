<%
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma", "No-Cache"

Dim Conexao, SQL, Tabela, Str, Exec, Mat, MSG

User = Trim(Session("Matricula"))
'Mat = Trim(request.querystring("Mat"))
Mat = Trim(Request("txtMatricula"))
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
			if (frmDados.txtMatricula.value == "") {
			alert("Preencha o campo login.");
			frmDados.txtMatricula.focus();
			return false;
		}
		if (frmDados.txtMatricula.value == "") {
			if (Valida(frmDados.txtMatricula.value) == false) {
				alert("O campo Matrícula não é válido.");
				frmDados.txtMatricula.focus();
				return false;
			}
		}
			if (frmDados.txtAtual.value == "") {
			alert("Preencha o campo Senha Atual.");
			frmDados.txtAtual.focus();
			return false;
		}
			if (frmDados.txtNova.value == "") {
			alert("Preencha o campo Nova Senha.");
			frmDados.txtNova.focus();
			return false;
		}
			if (frmDados.txtConfirma.value == "") {
			alert("Preencha o campo Confirmação Nova Senha.");
			frmDados.txtConfirma.focus();
			return false;
		}
			if (frmDados.txtConfirma.value !== frmDados.txtNova.value) {
			alert("Os campos Nova Senha e Confirmação Nova Senha não conferem.");
			frmDados.txtNova.focus();
			return false;
		}
		if (frmDados.txtEmail.value == "") {
			alert("Preencha o campo E-mail.");
			frmDados.txtEmail.focus();
			return false;
		}
		if (frmDados.txtOi.value == "") {
			alert("Preencha o campo Número Telefone.");
			frmDados.txtOi.focus();
			return false;
		}
				frmDados.action = "altera.asp?Mat="+frmDados.txtMatricula.value+"&Exec=S";
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
      <p align="center"><font face="Verdana" size="3"><b>Alterar Senha</b></font></td>
  </tr>
</table>
<form method="POST" name="frmDados">
<%
	If Trim(MSG) = "Sim" Then
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Você deve alterar agora sua senha, para sua segurança.');"
		Response.Write "</script>"
	End If
%>
<table border="0" width="100%">
  <tr>
      <td width="20%"><font face="Verdana" size="1"><b>Login:</b></font></td>
    <td width="80%">
        <p><input class="formfield" type="text" name="txtMatricula" value="<%=Session("usuario")%>" size="10" value=""></p>
    </td>
  </tr>
  <tr>
    <td width="20%"><font face="Verdana" size="1"><b>Senha Atual:</b></font></td>
    <td width="80%"><input class="formfield" type="password" name="txtAtual" size="10" value="1234" readonly></td>
  </tr>
  <tr>
    <td width="20%"><font face="Verdana" size="1"><b>Nova Senha:</b></font></td>
    <td width="80%"><input class="formfield" type="password" name="txtNova" size="10"></td>
  </tr>
  <tr>
    <td width="20%"><font face="Verdana" size="1"><b>Confirmação Nova Senha:</b></font></td>
    <td width="80%"><input class="formfield" type="password" name="txtConfirma" size="10"></td>
  </tr>
  <tr>
    <td width="20%"><font face="Verdana" size="1"><b>Digite seu e-mail:</b></font></td>
    <td width="80%"><input class="formfield" type="text" name="txtEmail" size="20">
    <b> @wedoservicos.com.br</b></td>
  </tr>
  <tr>
    <td width="20%"><font face="Verdana" size="1"><b>Digite seu Telefone:</b></font></td>
    <td width="80%"><input class="formfield" type="text" name="txtOi" size="20"><b> Ex.: DDD+Número</b></td>
  </tr>
  <tr>
    <td width="20%"><b><font face="Verdana" size="1">&nbsp;</font></b></td>
    <td width="80%"></td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
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
SQL = "SELECT * FROM SIS_Usuarios WHERE Usuario = '" & Mat & "'"
SQL = SQL & " --AND SENHA = '" & Trim(Request("txtAtual")) & "'"

Set Tabela = Conexao.execute(SQL)

If Tabela.EOF and Tabela.BOF Then
	Response.Write "<script language='javascript'>"
	Response.Write "alert('Sua senha foi alterada com sucesso! Efetue logon novamente!');"
	Response.Write "parent.parent.top.location.href='alteraOk.asp';"
	Response.Write "</script>"
	Response.End
Else
	SQL = "UPDATE SIS_Usuarios SET SENHA = '" & Trim(Request("txtNova")) & "', Tel = '" & Trim(Request("txtOi")) & "', Email = '" & Trim(Request("txtemail")) & "@wedoservicos.com.br' WHERE usuario = '" & Mat & "'"
	Set Tabela = Conexao.execute(SQL)
	If Err.number = 0 Then
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Sua senha foi alterada com sucesso! Efetue logon novamente!');"
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
