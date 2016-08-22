<%
Dim Tipo, Enviar, Excluir, CodigoE, ID, ResetSenha
Tipo = Request.QueryString("Tipo")
Enviar = Request.QueryString("Enviar")
Excluir = Request.QueryString("Excluir")
ResetSenha = Request.QueryString("ResetSenha")
ID = Request.QueryString("ID")
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Administra&ccedil;&atilde;o de Usuário</title>
<link rel="stylesheet" href="include/pgo.css" type="text/css">
<script language="javascript">
	function ValidaDados() {
		if (frmDados.txtNome.value == "") {
			alert("Preencha o campo Nome do Usuário.");
			frmDados.txtNome.focus();
			return false;
		}
		if (frmDados.txtTexto.value == "") {
			alert("Preencha o campo Login.");
			frmDados.txtTexto.focus();
			return false;
		}
		if (frmDados.txtSenha.value == "") {
			alert("Preencha o campo Senha.");
			frmDados.txtSenha.focus();
			return false;
		}
		if (frmDados.txtConf.value == "") {
			alert("Preencha o campo Confirmação de Senha.");
			frmDados.txtConf.focus();
			return false;
		}
		if (frmDados.txtAcesso.value == "0") {
			alert("Preencha o campo Acesso.");
			frmDados.txtAcesso.focus();
			return false;
		}
		if (frmDados.txtSupervisor.value == "0") {
			alert("Preencha o campo Acesso.");
			frmDados.txtSupervisor.focus();
			return false;
		}
		if (frmDados.txtSenha.value != frmDados.txtConf.value) {
			alert("Confirmação da Senha não Confere.");
			frmDados.txtConf.focus();
			return false;
		}
		if (confirm("Confirma a inclusão do Usuário?")) {
				frmDados.action = "usuarios.asp?Tipo="+frmDados.txtTipoJava.value+"&Enviar=S";
				frmDados.submit();
		}
		return false;
	}
</script>
</head>
<body>
<form method="POST" name="frmDados">
  <table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
    <tr bgcolor="#EBF3F1">
      <td colspan="2" bgcolor="#EBF3F1"> <p align="center"><b>Administra&ccedil;&atilde;o de Usuários</b></td>
    </tr>
    <input type="hidden" name="txtTipoJava" value="<%=Tipo%>">
    <!--#include file="include/AbreConexao.asp"-->
    <tr>
      <td>Nome do Usu&aacute;rio:</td>
      <td><input name="txtNome" type="text" class="formfield" size="50" maxlength="50"></td>
    </tr>
    <tr>
      <td width="24%"><i>Login:</i></td>
      <td width="76%"><input name="txtTexto" type="text" class="formfield" size="50" maxlength="50"></td>
    </tr>
    <tr>
      <td>Parametros de Senha:</td>
      <td>Senha:
           <input name="txtSenha" type="password" class="formfield" size="10" maxlength="10">
          Confirme a Senha:
           <input name="txtConf" type="password" class="formfield" size="10" maxlength="10">
        </td>
    </tr>
    <tr>
      <td>Acesso:</td>
      <td><select name="txtAcesso" type="text" class="textoform" size="1">
    <option value="0">Selecione Aqui</option>
    <option value="Administrador" >Administrador</option>
    <option value="Supervisor" >Supervisor</option>
    <option value="Auditoria" >Auditoria</option>
    <option value="Back-Office" >Back-Office</option>
    <option value="Vendas" >Vendas</option>
    </select></td>
    </tr>
    <tr>
      <td>Supervisor:</td>
      <td><select name="txtSupervisor" type="text" class="textoform" size="1">
    <option value="0">Selecione Aqui</option>
              <!--#include file="include/AbreConexao.asp"-->
              <%
'Consolidado de outros meses
SQL = " SELECT DISTINCT Nome as Nome FROM SIS_USUARIOS WHERE acesso in ('Supervisor', 'Administrador')"
SQL = SQL & " ORDER BY Nome "

Set RSBUSCA = server.createobject("ADODB.Recordset")
RSBUSCA.Open SQL, Conexao

Do While Not RSBUSCA.EOF

%>
              <option value="<%=RSBUSCA("Nome")%>"><%=RSBUSCA("Nome")%></option>
              <%
	RSBUSCA.MoveNext
Loop

	Set RSBUSCA = Nothing
	Set Conexao = Nothing
%>	
    </select></td>
    </tr>
    <tr>
      <td colspan="2"> <p align="left">
          <input name="Salvar" type="submit" onClick="return ValidaDados();" id="Salvar" value="Salvar" class="formbutton">
          &nbsp;
          <input class="formbutton" type="reset" value="Limpar" name="B2">
      </td>
    </tr>
  </table>
</form>
<%
Sub Inserir
%>
<!--#include file="include/AbreConexao.asp"-->
<%
	SQL = "INSERT INTO SIS_usuarios "
	SQL = SQL & "(usuario, Nome, Senha, acesso, Supervisor, Status) VALUES ('"& Request("txtTexto") &"', '"& Request("txtNome") &"', '"& Request("txtSenha") &"', '"& Request("txtAcesso") &"', '"& Request("txtSupervisor") &"', '1')"
Conexao.Execute(SQL)

If Enviar <> "" Then
		If Err.number = 0 Then
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Operação realizada com sucesso!\nEsperamos que o Usuário, "& Request("txtNome") &", Tenha um Bom Aproveitamento desta Ferramenta!');"
		Response.Write "location.href='usuarios.asp?Tipo="& Request("txtTipojava") &"';"
		Response.Write "</script>"
	Else
		If Err.number <> 0 Then
			Response.Write "<script language=""JavaScript"">alert(""Ocorreu um erro de execução\n\nErro:" & Err.number & "\n" & Err.description & _
				"\n\nInforme esse erro ao desenvolvedor"");</script>"
		End If
	End If
Else
		Response.Write "<script language=""JavaScript"">alert(""Alguma informação estava vazia. Favor preencher."");</script>"
End If
Set Conexao = Nothing
End Sub
If Enviar = "S"  Then
	Inserir
End If
%>
&nbsp;
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td colspan="13"> <p align="center"><b>Usuários Ativos</b></td>
  </tr>
  <tr bgcolor="#EBF3F1">
    <td width="5%" align="center"><i>Alterar</i></td>
    <td width="15%" align="center"> <p align="center"><i>Nome</i></td>
    <td width="15%" align="center"><div align="center">Usu&aacute;rio</div></td>
    <td width="5%" align="center"><div align="center">Acesso</div></td>
      </tr>
  <!--#include file="include/AbreConexao.asp"-->
  <%
SQL = "SELECT * FROM SIS_Usuarios WHERE usuario is not null and status = 1 ORDER BY Nome, Acesso, usuario "
Set RSBUSCAS = server.createobject("ADODB.Recordset")
RSBUSCAS.Open SQL, Conexao

Do While Not RSBUSCAS.EOF
%>
  <tr>
    <td width="5%" align="center"> 
<a href="usuarios.asp?Excluir=S&Tipo=<%=Tipo%>&ID=<%=RSBuscas("usuario")%>"><img alt="Excluir" border="0" src="imagens/cancelar.gif"></a>
<a href="usuarios.asp?ResetSenha=S&Tipo=<%=Tipo%>&ID=<%=RSBuscas("usuario")%>"><img width="16" height="16" alt="Reset Senha" border="0" src="imagens/cadeado.jpg"></a>
    </td>
    <td width="15%" align="left"><%=RSBuscas("Nome")%></td>
    <td width="15%" align="center"><%=RSBuscas("usuario")%> <div align="center"></div></td>
    <td width="5%" align="center"><%=RSBuscas("acesso")%> <div align="center"></div></td>
      </tr>
  <%
	RSBUSCAS.Movenext
Loop
%>
</table>
<%
If Excluir = "S" Then
SQL  = "UPDATE SIS_USUARIOS SET STATUS = 0 WHERE usuario = '" & ID & "'"
	Conexao.Execute(SQL)
		Response.Write "<script language='javascript'>"
     	Response.Write "location.href='usuarios.asp?Tipo="&Request.querystring("Tipo")&"';"
		Response.Write "alert('O Login " & ID & " esta INATIVO com Sucesso.\nObrigado Por Utilizar Esta Ferramenta!');"
	    Response.Write "</script>"
End If
%>

<%
If ResetSenha = "S" Then
SQL  = "UPDATE SIS_USUARIOS SET SENHA = 123 WHERE usuario = '" & ID & "'"
	Conexao.Execute(SQL)
		Response.Write "<script language='javascript'>"
     	Response.Write "location.href='usuarios.asp?Tipo="&Request.querystring("Tipo")&"';"
		Response.Write "alert('O reset da senha do Login " & ID & " foi realizado com sucesso\nObrigado Por Utilizar Esta Ferramenta!');"
	    Response.Write "</script>"
End If
%>
</body>
</html>