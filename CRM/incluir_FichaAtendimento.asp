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

Dim Obs, Atualiza, Enviar, Observa, Data, ID, NomeCliente, DDDTelefone, Endereco, Numero, Complemento, Bairro, CEP, Cidade, Estado, Email, DDDTelefoneCEL, ResultadoChamada

ID = request.querystring("ID")
IDVisita = request.querystring("IDVisita")
IDVendedor = request.querystring("IDVendedor")
Enviar = request.querystring("Enviar")
Atualiza = request.querystring("Atualiza")
Data = Request.QueryString("Dia")
ResultadoChamada = request.querystring("ResultadoChamada")
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
<script language="javascript">
	function ValidaDados() {
		if (frmDados.txtResultado.value == "0") {
			alert("Escolha uma opção no campo Resultado da Chamada.");
			frmDados.txtResultado.focus();
			return false;
		}
		if (frmDados.txtDetResultado.value == "0") {
			alert("Escolha uma opção no campo Detalhe do Resultado.");
			frmDados.txtDetresultado.focus();
			return false;
		}
		if (confirm("Confirma a Ficha de Atendimento?")) {
				frmDados.action = "incluir_FichaAtendimento.asp?ID="+frmDados.txtID.value+"&ResultadoChamada="+frmDados.txtResultadoChamada.value+"&Atualiza=S";
				frmDados.submit();
		}
		return false;
	}
</script>
</head>
<form method="POST" name="frmDados">
<%
Sub Localizar
%>
<!--#include file="AbreConexao.asp"-->
<%

SQLL = " SELECT * from tbProspects "
SQLL = SQLL & " WHERE ID = " & ID & " "

Set RSBUSCA = server.createobject("ADODB.Recordset")
RSBUSCA.Open SQLL, Conexao

If Not RSBUSCA.EOF Then

	ID = RSBUSCA("ID")
	NomeCliente = RSBUSCA("NomeCliente")
	DDDTelefone = RSBUSCA("DDDTelefone")
	Endereco = RSBUSCA("Endereco")
	Numero = RSBUSCA("Numero")
	Complemento = RSBUSCA("Complemento")		
	Bairro = RSBUSCA("Bairro")
	CEP = RSBUSCA("CEP")	
	Cidade = RSBUSCA("Cidade")
	Estado = RSBUSCA("Estado")
	Email = RSBUSCA("email")
	DDDTelefoneCEL = RSBUSCA("DDDTelefoneCEL")
	Obs = RSBUSCA("Observacao")


End If
Conexao.Close
Set SQL = Nothing
Set RSBUSCA = Nothing
End Sub

If ID <> "" and Enviar <> "S" Then
	Localizar
End If
%>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="PreenchetxtData('SpanData', 'txtData');">
<hr color="#002b4c" size="1">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
  <tr>
    <td width="100%" colspan="2">
      <p align="center"><font face="Verdana" size="3"><b>Ficha de Atendimento - PIZZA HUT</b></font></td>
  </tr>
</table>
<br>

<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td colspan="4" width="50%" align="center">Nome Cliente</td>
    <td colspan="3" width="50%" align="center">DDD e Telefone</td>
  </tr>
  <tr>
    <td colspan="4" width="50%" align="center"><p align="center"><%=NomeCliente%></td>
    <td colspan="3" width="50%" align="center"><p align="center"><%=DDDTelefone%></td>
  </tr>

  <tr bgcolor="#EBF3F1">
    <td width="16%" align="center">CEP</td>
    <td width="16%" align="center">Endere&ccedil;o</td>
    <td width="16%" align="center">Bairro</td>
    <td width="16%" align="center">Cidade</td>
    <td width="16%" align="center">E-mail</td>
    <td width="16%" align="center">Telefone Celular</td>
  </tr>
  <tr>
    <td width="16%" align="center"><b><%=CEP%></b></td>
    <td width="16%" align="center"><%=Endereco%></td>
    <td width="16%" align="center"><%=Bairro%></td>
    <td width="16%" align="center"><%=Cidade%></td>
    <td width="16%" align="center"><%=Email%></td>
    <td width="16%" align="center"><%=DDDTelefoneCEL%></td>
  </tr>


</table>
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td width="33%" align="center"><p align="center"><b>Resultado da Chamada</b></td>
    <td width="33%" align="center"><p align="center"><b>Detalhe do Resultado</b></td>
  </tr>
  <tr>
    <td align="left">  <p align="center">

<select name="txtResultado" type="text" class="textoform" size="1" onChange="location=this.options[this.selectedIndex].value;mostra_destaque(emdestaque)">
    <option value="0" selected>Selecione aqui</option>
<option value="incluir_FichaAtendimento.asp?ID=<%=ID%>&ResultadoChamada=Pedido"<% If ResultadoChamada = "Pedido" Then %> selected <%End If%>>Pedido</option>
<option value="incluir_FichaAtendimento.asp?ID=<%=ID%>&ResultadoChamada=Informação"<% If ResultadoChamada = "Informação" Then %> selected <%End If%>>Informação</option>
<option value="incluir_FichaAtendimento.asp?ID=<%=ID%>&ResultadoChamada=Sugestão"<% If ResultadoChamada = "Sugestão" Then %> selected <%End If%>>Sugestão</option>
<option value="incluir_FichaAtendimento.asp?ID=<%=ID%>&ResultadoChamada=Reclamação"<% If ResultadoChamada = "Reclamação" Then %> selected <%End If%>>Reclamação</option>
<option value="incluir_FichaAtendimento.asp?ID=<%=ID%>&ResultadoChamada=Chamada Nula"<% If ResultadoChamada = "Chamada Nula" Then %> selected <%End If%>>Chamada Nula</option>
   </select>    
    </td>
    <td align="left">  <p align="center">
    
    <select name="txtDetResultado" type="text" class="textoform" size="1">
    <option value="0">Selecione aqui</option>
<%
Sub Combo
%>
    <!--#include file="AbreConexao.asp"-->
    <%
	SQLCombo = "SELECT DISTINCT DetResultado FROM SIS_ResultadoChamadas WHERE Resultado = '" & ResultadoChamada & "' and Status = 1 ORDER BY DetResultado ASC "
	Set RSBUSCACombo = server.createobject("ADODB.Recordset")
	RSBUSCACombo.Open SQLCombo, Conexao

	Do While Not RSBUSCACombo.EOF
%>
    <option value="<%=RSBUSCACombo("DetResultado")%>"><%=RSBUSCACombo("DetResultado")%></option>
    <%
		RSBUSCACombo.MoveNext
	Loop
RSBUSCACombo.Close
Set SQLCombo = Nothing
Set Conexao = Nothing
End Sub
Combo
%>
    </select>
    </td>
  </table>
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">

  <tr bgcolor="#EBF3F1">
    <td width="100%" align="center"><p align="center"><b>Código Promocional:</b></td>
  </tr>
    <tr>
    <td width="100%" ><p align="center">
        <input name="txtCodPromo" type="text" class="formfield" size="10" maxlength="3"></p>
    </td>
  </tr>

  <tr bgcolor="#EBF3F1">
    <td width="100%" align="center"><p align="center"><b>Observações:</b></td>
  </tr>
    <tr>
    <td width="100%" ><p align="center">
        <textarea class="formfield" rows="4" name="txtObs" cols="102" maxlength="250"><%=TRIM(Obs)%></textarea></p>
    </td>
  </tr>
</table>
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
    <tr>
    <td width="100%" align="center">
    <input type="hidden" name="txtID" value="<%=ID%>">
    <input type="hidden" name="txtResultadoChamada" value="<%=ResultadoChamada%>">
    <input name="Salvar" type="submit" onClick="return ValidaDados();" id="Salvar" value="Salvar" class="formbutton">&nbsp;<input class="formbutton" type="reset" value="Limpar" name="B2">
    </td>
  </tr>
</table>
&nbsp;
</form>
<%
Sub Atualizar
Hoje = Year(Now)&"-"&Month(Now)&"-"&Day(Now)&" "&TimeValue(Now)
Data = Year(Data)&"-"&Month(Data)&"-"&Day(Data)
Observa = Request("txtObs")
Observa = Replace(Observa, Chr(39), Chr(95))
If IsNull(Observa) or Observa = "" Then
	Observa = "Sem Observação"
End If
%>
<!--#include file="AbreConexao.asp"-->
<%
	SQL4 = "INSERT INTO tbFichasAtendimento "
	SQL4 = SQL4 & "(IDProspects, DataCriacao, ResultadoChamada, DetResultado, ResponsavelCriacao, Observacao, CodPromo) VALUES ('"& Request("txtID") &"', getdate(), '"& ResultadoChamada &"', '"& Request("txtDetResultado") &"', '"&User&"', '"& Observa &"', '"& Request("txtCodPromo") &"')"
'response.write SQL4
Conexao.Execute(SQL4)

	If Err.number = 0 Then
		Response.Write "<script language='javascript'>"
If Request("txtDetResultado") = "Pedido Realizado" Then
		Response.Write "alert('Pedido PIZZA HUT Realizado com Sucesso!');"
		Response.Write "location.href='verificacao_telefonia.asp';"
Else
		Response.Write "alert('O Resultado para a PIZZA HUT foi atualizado com sucesso.');"
		Response.Write "location.href='verificacao_telefonia.asp';"
End If
		Response.Write "</script>"
	Else
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Ocorreu um erro no envio das informacões. Favor entrar em contato com o Administrador do Sistema.');"
		Response.Write "</script>"
		Response.End
	End If
	Conexao.Close
	Set SQL = Nothing
End Sub

If Request("txtID") <> "" and Enviar <> "S" Then
	Atualizar
End If
%>

<%
Sub Historico
%>
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td colspan="10"> <p align="center"><b>Histórico Visitas</b></td>
  </tr>
  <tr bgcolor="#EBF3F1">
    <td width="5%" align="center"> <p align="center"><i>DDDTelefone</i></td>
    <td width="5%" align="center"> <p align="center"><i>Nome Cliente</i></td>
    <td width="5%" align="center"> <p align="center"><i>Resultado Chamada</i></td>
    <td width="5%" align="center"><p align="center"><i>Detalhe Resultado</i></td>
    <td width="5%" align="center"><p align="center"><i>Data</i></td>
    <td width="20%" align="center"><p align="center"><i>Observação</i></td>
     </tr>

<!--#include file="AbreConexao.asp"-->
<%

SQL = " SELECT     P.DDDTelefone, P.NomeCliente, FA.DataCriacao, FA.ResultadoChamada, FA.DetResultado, FA.Observacao "
SQL = SQL & " FROM tbFichasAtendimento AS FA INNER JOIN "
SQL = SQL & " tbProspects AS P ON FA.IDProspects = P.ID "
SQL = SQL & " where FA.IDProspects = "& ID &" "
SQL = SQL & "ORDER BY FA.DataCriacao ASC "
Set RSBUSCAS = server.createobject("ADODB.Recordset")
'RESPONSE.WRITE SQL
RSBUSCAS.Open SQL, Conexao

Do While Not RSBUSCAS.EOF
%>
  <tr>
    <td width="5%" align="center"><p align="center"><%=RSBuscas("DDDTelefone")%>&nbsp;</td>
    <td width="5%" align="center"><p align="center"><%=RSBuscas("NomeCliente")%>&nbsp;</td>
    <td width="5%" align="center"><p align="center"><%=RSBuscas("ResultadoChamada")%>&nbsp;</td>
    <td width="5%" align="center"><p align="center"><%=RSBuscas("DetResultado")%>&nbsp;</td>
    <td width="5%" align="center"><p align="center"><%=RSBuscas("DataCriacao")%>&nbsp;</td>
    <td width="20%" align="center"><%=RSBuscas("Observacao")%>&nbsp;</td>
      </tr>
  <%
	RSBUSCAS.Movenext
Loop

End Sub
If ID <> "" and Enviar <> "S" Then
	Historico
End If
%>

</table>

</body>

