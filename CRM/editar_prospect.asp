<%
Dim NomeFantasia, RazaoSocial, CNPJ, Bairro, Cidade, CEP, Contato, Telefone, Vendedor, Revenda, Consultor
Dim DataCadastro, DataPrimeiraVisita, DataUltimaVisita, UltimoFunil, UltimoDetFunil, ProximaVisita, ProdutoCliente, Operadora, Tempo
Dim Data, Enviar, ID, IDVendedor, IDRevenda, IDConsultor
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
<title>Administra&ccedil;&atilde;o de Usuário</title>
<link rel="stylesheet" href="../include/pgo.css" type="text/css">
<script language="JavaScript" src="../include/tabulacaoautomatica.js"></script>
<script language="javascript">
	function ValidaDados() {
		if (frmDados.txtVendedor.value == "0") {
			alert("Preencha o campo Nome do Vendedor.");
			frmDados.txtVendedor.focus();
			return false;
		}
		if (frmDados.txtEmpresa.value == "0") {
			alert("Preencha o campo Nome do Agente.");
			frmDados.txtEmpresa.focus();
			return false;
		}
		if (frmDados.txtConsultor.value == "0") {
			alert("Preencha o campo Nome do Consultor.");
			frmDados.txtConsultor.focus();
			return false;
		}				
		if (frmDados.txtDocumento.value == "") {
			alert("Preencha o campo CPF/CNPJ.");
			frmDados.txtDocumento.focus();
			return false;
		}
		if (frmDados.txtCEP.value == "") {
			alert("Preencha o campo CEP.");
			frmDados.txtCEP.focus();
			return false;
		}
		if (frmDados.txtBairro.value == "0") {
			alert("Preencha o campo Bairro.");
			frmDados.txtBairro.focus();
			return false;
		}
		if (frmDados.txtCidade.value == "0") {
			alert("Preencha o campo Cidade.");
			frmDados.txtCidade.focus();
			return false;
		}
		if (frmDados.txtTelefone.value == "") {
			alert("Preencha o campo Telefone.");
			frmDados.txtTelefone.focus();
			return false;
		}								
		if (confirm("Confirma a atualização do cliente?")) {
				frmDados.action = "editar_prospect.asp?Enviar=S";
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
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<form method="POST" name="frmDados">
<%
Sub Localizar
%>
<!--#include file="AbreConexao.asp"-->
<%

SQLL = " SELECT P.*, V.*, R.*, C.*, DateDiff(Day, P.DataUltimaVisita, getdate()) as Tempo "
SQLL = SQLL & "  FROM tbProspects AS P INNER JOIN "
SQLL = SQLL & "  tbVendedor AS V ON P.IDVendedor = V.IDVendedor INNER JOIN "
SQLL = SQLL & "  tbRevenda AS R ON V.IDRevenda = R.IDRevenda INNER JOIN "
SQLL = SQLL & "  tbConsultor AS C ON R.IDConsultor = C.IDConsultor "
SQLL = SQLL & " WHERE (P.ID = " & ID & ") "

Set RSBUSCA = server.createobject("ADODB.Recordset")
RSBUSCA.Open SQLL, Conexao

If Not RSBUSCA.EOF Then

	ID = RSBUSCA("ID")
	NomeFantasia = RSBUSCA("NomeFantasia")
	RazaoSocial = RSBUSCA("RazaoSocial")
	CNPJ = RSBUSCA("CNPJ")
	Bairro = RSBUSCA("Bairro")
	Cidade = RSBUSCA("Cidade")
	CEP = RSBUSCA("CEP")
	Contato = RSBUSCA("Contato")
	Telefone = RSBUSCA("Telefone")
	DataPrimeiraVisita = RSBUSCA("DataPrimeiraVisita")

	Vendedor = RSBUSCA("Vendedor")
	Revenda = RSBUSCA("Revenda")
	Consultor = RSBUSCA("Consultor")
	DataCadastro = RSBUSCA("DataCadastro")
	DataUltimaVisita = RSBUSCA("DataUltimaVisita")
	UltimoFunil = RSBUSCA("UltimoFunil")
	UltimoDetFunil = RSBUSCA("UltimoDetFunil")
	ProximaVisita = RSBUSCA("ProximaVisita")
	Tempo = RSBUSCA("Tempo")
	IDVendedor = RSBUSCA("IDVendedor")
	IDRevenda = RSBUSCA("IDRevenda")	
	IDConsultor = RSBUSCA("IDConsultor")	

	ProdutoCliente = RSBUSCA("ProdutoCliente")
	Operadora = RSBUSCA("Operadora")
	Voz1 = RSBUSCA("Voz1")
	Voz2 = RSBUSCA("Voz2")
	Dados1 = RSBUSCA("Dados1")
	Dados2 = RSBUSCA("Dados2")
	Velocidade1 = RSBUSCA("Velocidade1")
	Velocidade2 = RSBUSCA("Velocidade2")


End If
Conexao.Close
Set SQL = Nothing
Set RSBUSCA = Nothing
End Sub

If ID <> "" and Enviar <> "S" Then
	Localizar
End If
%>

  <table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
    <tr bgcolor="#EBF3F1">
      <td colspan="2" bgcolor="#EBF3F1"> <p align="center"><b>Editar Prospects</b></td>
    </tr>
    <!--#include file="AbreConexao.asp"-->
    <tr>
      <td>Nome do Vendedor:<%=Vendedor%></td>
      <td>  <select name="txtVendedor" type="text" class="textoform" size="1">
    <option value="0">Selecione aqui</option>
    <%
Sub ComboVendedor
%>
    <!--#include file="AbreConexao.asp"-->
    <%
If Empresa <> 1 Then
	SQLCombo = "SELECT * FROM tbVendedor where IDRevenda = "&Empresa&" and Status = 1 order by Vendedor"
Else	
	SQLCombo = "SELECT * FROM tbVendedor where Status = 1 order by Vendedor"	
End If	
	Set RSBUSCACombo = server.createobject("ADODB.Recordset")
	RSBUSCACombo.Open SQLCombo, Conexao

	Do While Not RSBUSCACombo.EOF
%>
    <option value="<%=RSBUSCACombo("IDVendedor")%>" <% If IDVendedor = RSBUSCACombo("IDVendedor") Then %>selected <%End If%>><%=RSBUSCACombo("Vendedor")%></option>
    <%
		RSBUSCACombo.MoveNext
	Loop
RSBUSCACombo.Close
Set SQLCombo = Nothing
Set Conexao = Nothing
End Sub
ComboVendedor
%>
    </select></td>
    </tr>
    <tr>
      <td>Empresa:</td>
      <td><select name="txtEmpresa" type="text" class="textoform" size="1">
    <option value="0">Selecione aqui</option>
    <%
Sub ComboEmpresa
%>
    <!--#include file="AbreConexao.asp"-->
    <%

If Empresa <> 1 Then
	SQLCombo = "SELECT * FROM tbRevenda where IDRevenda = "&Empresa&" and Status = 1 order by Revenda"
Else
	SQLCombo = "SELECT * FROM tbRevenda where Status = 1 order by Revenda"
End If

	Set RSBUSCACombo = server.createobject("ADODB.Recordset")
	RSBUSCACombo.Open SQLCombo, Conexao

	Do While Not RSBUSCACombo.EOF
%>
    <option value="<%=RSBUSCACombo("IDRevenda")%>"  <% If IDRevenda = RSBUSCACombo("IDRevenda") Then %>selected <%End If%>><%=RSBUSCACombo("Revenda")%></option>
    <%
		RSBUSCACombo.MoveNext
	Loop
RSBUSCACombo.Close
Set SQLCombo = Nothing
Set Conexao = Nothing
End Sub
ComboEmpresa
%>
    </select></td>
    </tr>
    <tr>
      <td>Consultor:</td>
      <td><select name="txtConsultor" type="text" class="textoform" size="1">
    <option value="0">Selecione aqui</option>
    <%
Sub ComboConsultor
%>
    <!--#include file="AbreConexao.asp"-->
    <%

If Empresa <> 1 Then
	SQLCombo = " SELECT C.IDCOnsultor, C.Consultor, R.IDRevenda FROM tbRevenda AS R INNER JOIN tbConsultor AS C ON R.IDConsultor = C.IDConsultor WHERE (R.IDRevenda = "&Empresa&") "
Else
	SQLCombo = "SELECT C.IDCOnsultor, C.Consultor FROM tbRevenda AS R INNER JOIN tbConsultor AS C ON R.IDConsultor = C.IDConsultor GROUP BY C.IDCOnsultor, C.Consultor"
End If

	Set RSBUSCACombo = server.createobject("ADODB.Recordset")
	RSBUSCACombo.Open SQLCombo, Conexao

	Do While Not RSBUSCACombo.EOF
%>
    <option value="<%=RSBUSCACombo("IDConsultor")%>"  <% If IDConsultor = RSBUSCACombo("IDConsultor") Then %>selected <%End If%>><%=RSBUSCACombo("Consultor")%></option>
    <%
		RSBUSCACombo.MoveNext
	Loop
RSBUSCACombo.Close
Set SQLCombo = Nothing
Set Conexao = Nothing
End Sub
ComboConsultor
%>
    </select></td>
    </tr>
    <tr>
      <td width="24%"><i>CPF/CNPJ:</i></td>
      <td width="76%"> <input name="txtDocumento" type="text" class="formfield" onKeyPress="return Valida(event);" onKeyUp="return autoTab(this, 15, event);" value="<%=CNPJ%>" size="50" maxlength="50" readonly="readonly"></td>
    </tr>
    <tr>
      <td>Nome Fantasia:</td>
      <td><input name="txtNomeFantasia" type="text" class="formfield" size="50" maxlength="50" value="<%=NomeFantasia%>"></td>
    </tr>
    <tr>
      <td>Razão Social:</td>
      <td><input name="txtRazaoSocial" type="text" class="formfield" size="50" maxlength="50" value="<%=RazaoSocial%>"></td>
    </tr>
    <tr>
      <td>CEP:</td>
      <td><input name="txtCEP" type="text" class="formfield" size="50" maxlength="50" onKeyUp="return autoTab(this, 8, event);" onKeyPress="return Valida(event);" value="<%=CEP%>"></td>
    </tr>    
    <tr>
      <td>Bairro:</td>
      <td><select name="txtBairro" type="text" class="textoform" size="1">
    <option value="0">Selecione aqui</option>
    <option value="Aldeota" <% If Bairro = "Aldeota" Then %>selected <%End If%>>Aldeota</option>
    <option value="Centro"  <% If Bairro = "Centro" Then %>selected <%End If%>>Centro</option>    
    </select></td>
    </tr>
    <tr>
      <td>Cidade:</td>
      <td><select name="txtCidade" type="text" class="textoform" size="1">
    <option value="0">Selecione aqui</option>
    <option value="Fortaleza" <% If Cidade = "Fortaleza" Then %>selected <%End If%> >Fortaleza</option>
    <option value="Caucaia" <% If Cidade = "Caucaia" Then %>selected <%End If%> >Caucaia</option>    
    <option value="Maracanau" <% If Cidade = "Maracanau" Then %>selected <%End If%> >Maracanau</option>        
    </select></td>
    </tr>

    <tr>
      <td>Contato:</td>
      <td><input name="txtContato" type="text" class="formfield" size="50" maxlength="50" value="<%=Contato%>"></td>
    </tr>    
    <tr>
      <td>Telefone:</td>
      <td><input name="txtTelefone" type="text" class="formfield" size="50" maxlength="50" onKeyUp="return autoTab(this, 10, event);" onKeyPress="return Valida(event);" value="<%=Telefone%>"></td>
    </tr>
    <tr>
      <td>Produto do Cliente:</td>
      <td><input name="txtProdutoCliente" type="text" class="formfield" size="50" maxlength="50" value="<%=ProdutoCliente%>"></td>
    </tr>
    <tr>
      <td>Operadora Atual:</td>
      <td><select name="txtOperadora" type="text" class="textoform" size="1">
    <option value="0">Selecione aqui</option>
    <option value="Oi" <% If Operadora = "Oi" Then %>selected <%End If%>>Oi</option>
    <option value="Embratel" <% If Operadora = "Embratel" Then %>selected <%End If%>>Embratel</option>
    <option value="Intelig" <% If Operadora = "Intelig" Then %>selected <%End If%>>Intelig</option>
    </select></td>
    </tr>
<input type="hidden" name="txtID" value="<%=ID%>">        
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

Data = Year(Data)&"-"&Month(Data)&"-"&Day(Data)

    SQL = " UPDATE tbProspects SET "
	SQL = SQL & " IDVendedor = '"& Request("txtVendedor") &"', IDRevenda = '"& Request("txtEmpresa") &"', "
	SQL = SQL & " IDConsultor = '"& Request("txtConsultor") &"', CNPJ = '"& Request("txtDocumento") &"', "
	SQL = SQL & " NomeFantasia = '"& Request("txtNomeFantasia") &"', RazaoSocial = '"& Request("txtRazaoSocial") &"', "
	SQL = SQL & " CEP = '"& Request("txtCEP") &"', Bairro = '"& Request("txtBairro") &"', Cidade = '"& Request("txtCidade") &"', "
	SQL = SQL & " Contato = '"& Request("txtContato") &"', Telefone = '"& Request("txtTelefone") &"', "
	SQL = SQL & " Operadora = '"& Request("txtOperadora") &"', ProdutoCliente = '"& Request("txtProdutoCliente") &"'"
	SQL = SQL & " WHERE ID = '"& Request("txtID") &"' "
	Conexao.Execute(SQL)

If Enviar <> "" Then
		If Err.number = 0 Then
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Operação realizada com sucesso!');"
		Response.Write "location.href='verificacao_prospects.asp?Enviar=S';"
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
</body>
</html>