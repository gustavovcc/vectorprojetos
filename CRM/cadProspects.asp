<%
Dim Data, Enviar, ID, IDVendedor, IDcadFichaOi
Enviar = Request.QueryString("Enviar")
ID = Request.QueryString("ID")
Data = Request.QueryString("Dia")
Vendedor = Request.QueryString("IDVendedor")
Empresa = Trim(Session("Empresa"))
User = Trim(Session("usuario"))

If User = "" Then
	Response.Write "<script language='javascript'>"
	Response.Write "alert('Efetue o logon novamente.');"
	Response.Write "parent.parent.top.location.href='../default.asp';"
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
<title>WeDo</title>
<link rel="stylesheet" href="../include/pgo.css" type="text/css">
<script language="javascript" src="../Include/MostraCalendario.js"></script>
<script language="JavaScript" src="../include/tabulacaoautomatica.js"></script>
<script language="javascript">
	function ValidaDados() {
		if (frmDados.txtNomeCliente.value == "") {
			alert("Preencha o campo Nome do Cliente.");
			frmDados.txtNomeCliente.focus();
			return false;
		}
		if (frmDados.txtTelefoneFixo.value == "") {
			alert("Preencha o campo Telefone Fixo.");
			frmDados.txtTelefoneFixo.focus();
			return false;
		}
		if (frmDados.txtCidade.value == "0") {
			alert("Preencha o campo Cidade.");
			frmDados.txtCidade.focus();
			return false;
		}
		if (frmDados.txtEstado.value == "0") {
			alert("Preencha o campo Estado.");
			frmDados.txtEstado.focus();
			return false;
		}
		if (confirm("Confirma a inclusão do cadastro?")) {
				frmDados.action = "cadProspects.asp?Enviar=S";
				frmDados.submit();
		}
		return false;
	}
</script>
<script language="javascript">
function Valida(evento){
  var tecla = evento.keyCode;
  return (((tecla>=48)&&(tecla<=57)) || (tecla==45));
}
</script>

</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form method="POST" name="frmDados">
  <table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
    <tr bgcolor="#EBF3F1">
      <td colspan="2" bgcolor="#EBF3F1"> <p align="center"><b>Cadastro Novo Cliente</b></td>
    </tr>
    <tr>
      <td width="24%">Nome do Vendedor:</td>
      <td width="76%">  <%=User%>
	</td>
    </tr>
    <tr>
      <td>Nome Cliente*:</td>
      <td><input name="txtNomeCliente" type="text" class="formfield" size="50" maxlength="50"></td>
    </tr>
    <tr>
      <td>Telefone Fixo*:</td>
      <td><input name="txtTelefoneFixo" type="text" class="formfield" onKeyPress="return Valida(event);" onKeyUp="return autoTab(this, 10, event);" value="<%=Request.QueryString("DDDTelefone")%>" size="50" maxlength="50"></td>
    </tr>
    <tr>
      <td>CEP:</td>
      <td><input name="txtCEP" type="text" class="formfield" size="50" maxlength="50" onKeyUp="return autoTab(this, 8, event);" onKeyPress="return Valida(event);"></td>
    </tr>
    <tr>
      <td>E-mail:</td>
      <td><input name="txtEmail" type="text" class="formfield" size="50" maxlength="50"></td>
    </tr>
    <tr>
      <td>Telefone Celular:</td>
      <td><input name="txtTelefoneCelular" type="text" class="formfield" size="50" maxlength="50" onKeyUp="return autoTab(this, 10, event);" onKeyPress="return Valida(event);"></td>
    </tr>
    <tr>
      <td>Endereço:</td>
      <td><input name="txtEndereco" type="text" class="formfield" size="50" maxlength="50"></td>
    </tr>
    <tr>
      <td>Número:</td>
      <td><input name="txtNumero" type="text" class="formfield" size="50" maxlength="50"></td>
    </tr>
    <tr>
      <td>Complemento:</td>
      <td><input name="txtComplemento" type="text" class="formfield" size="50" maxlength="50"></td>
    </tr>
    <tr>
      <td>Bairro:</td>
      <td><input name="txtBairro" type="text" class="formfield" size="50" maxlength="50"></td>
    </tr>
    <tr>
      <td>Cidade:</td>
      <td><select name="txtCidade" type="text" class="textoform" size="1">
    <option value="0">Selecione aqui</option>
    <option value="Fortaleza" selected>Fortaleza</option>
    <option value="João Pessoa">João Pessoa</option>
    </select></td>
    </tr>
    <tr>
      <td>Estado:</td>
      <td><select name="txtEstado" type="text" class="textoform" size="1">
    <option value="0">Selecione aqui</option>
    <option value="Ceara" selected>Ceará</option>
    <option value="Paraiba">Paraiba</option>
    </select></td>
    </tr>

    <tr>
      <td>Observação:</td>
      <td>
	  <textarea class="formfield" rows="4" name="txtObserva" cols="102" maxlength="250"></textarea>
      </td>
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
<!--#include file="AbreConexao.asp"-->
<%

SQL1 = "SELECT ID, DDDTelefone FROM tbProspects WHERE DDDTelefone = '"& Request("txtTelefoneFixo") &"' "
Set RSBUSCAS = server.createobject("ADODB.Recordset")
RSBUSCAS.Open SQL1, Conexao

Do While Not RSBUSCAS.EOF

DDDTelefone = RSBuscas("DDDTelefone")
IDProspect = RSBuscas("ID")

	RSBUSCAS.Movenext
Loop

If DDDTelefone = "" Then

Hoje1 = Year(Now)&"-"&Month(Now)&"-"&Day(Now)&" "&TimeValue(Now)
	SQL = "INSERT INTO tbProspects "
	SQL = SQL & "(NomeCliente, CEP, Endereco, Numero, Complemento, Bairro, Cidade, Estado, Email, DDDTelefone, DDDTelefoneCEL, Observacao, DataCriacao, ResponsavelCriacao) VALUES ('"& Request("txtNomeCliente") &"', '"& Request("txtCEP") &"', '"& Request("txtEndereco") &"', '"& Request("txtNumero") &"', '"& Request("txtComplemento") &"', '"& Request("txtBairro") &"', '"& Request("txtCidade") &"', '"& Request("txtEstado") &"', '"& Request("txtEmail") &"', '"& Request("txtTelefoneFixo") &"', '"& Request("txtTelefoneCelular") &"', '"& Request("txtObserva") &"', CONVERT(DATETIME, '"&Hoje1&"', 102), '"&User&"' )"
'response.write SQL
Conexao.Execute(SQL)

SQLL = "SELECT max(ID) as ID FROM tbProspects "
SQLL = SQLL & " Where ResponsavelCriacao = '"&User&"' "
'response.write SQLL
Set RSBUSCA = server.createobject("ADODB.Recordset")
RSBUSCA.Open SQLL, Conexao
If Not RSBUSCA.EOF Then
	IDcadFichaOi = RSBUSCA("ID")
End If


If Enviar <> "" Then
		If Err.number = 0 Then
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Cadastro Realizado com Sucesso!! ID Cliente: "&IDcadFichaOi&" \nEncaminhado para o Cadastro da Ficha de Atendimento.');"
		Response.Write "location.href='incluir_FichaAtendimento.asp?User="&User&"&ID="&IDcadFichaOi&"';"
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

Else

		Response.Write "<script language='javascript'>"
		Response.Write "alert('Cliente JÁ EXISTE!\nOperação NÃO realizada com sucesso!');"
		Response.Write "location.href='incluir_FichaAtendimento.asp?User="&User&"&ID="&IDProspect&"';"
		Response.Write "</script>"

End If

End Sub
If Enviar = "S"  Then
	Inserir
End If
%>
</body>
</html>