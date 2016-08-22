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

Dim Obs, Atualiza, Enviar, Observa, Data, ID, NomeCliente, DDDTelefone, Valor, DataPedido, Filial, Loja, ResultadoChamada

ID = request.querystring("ID")
Loja = request.querystring("Loja")
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
		if (frmDados.txtCallCenter.value == "0") {
			alert("Escolha uma opção no campo Call Center.");
			frmDados.txtCallCenter.focus();
			return false;
		}
		if (frmDados.txtMotoBoy.value == "0") {
			alert("Escolha uma opção no campo Moto Boy.");
			frmDados.txtMotoBoy.focus();
			return false;
		}
		if (frmDados.txtTempo.value == "0") {
			alert("Escolha uma opção no campo Tempo de Entrega.");
			frmDados.txtTempo.focus();
			return false;
		}
		if (frmDados.txtProduto.value == "0") {
			alert("Escolha uma opção no campo Produto.");
			frmDados.txtProduto.focus();
			return false;
		}		
		if (confirm("Confirma a Ficha de Atendimento?")) {
				frmDados.action = "incluir_FichaAtendimento_PosVenda.asp?DDDTelefone="+frmDados.txtDDDTelefone.value+"&NomeCliente="+frmDados.txtNomeCliente.value+"&Atualiza=S";
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

If Loja = "Pastel na Hora" Then

SQLL = " SELECT * FROM ( "
SQLL = SQLL & " SELECT VEN_NRTELE AS TELEFONE, VEN_NMCLIE AS NOMECLIENTE, VEN_DHRECE AS DATA, VEN_VLRNOT AS VALOR, "
SQLL = SQLL & " CASE WHEN VEN_NRLOJA = '3' THEN 'O. Paiva' ELSE  "
SQLL = SQLL & " CASE WHEN VEN_NRLOJA = '5' THEN 'Maraponga' ELSE  "
SQLL = SQLL & " CASE WHEN VEN_NRLOJA = '4' THEN '' ELSE 'Outros' END END END as FILIAL, 'Pastel na Hora' as LOJA "
SQLL = SQLL & " FROM PastelNAHORA_DELIVERY..TB_MOVVENDA "
SQLL = SQLL & " WHERE VEN_TPVEND = 2 AND VEN_STATUS = 0 and  "
SQLL = SQLL & " day(VEN_DHRECE) = day(getdate()-1) and  "
SQLL = SQLL & " Month(VEN_DHRECE) = Month(getdate()) and  "
SQLL = SQLL & " Year(VEN_DHRECE) = Year(getdate())  "
SQLL = SQLL & " ) A "
SQLL = SQLL & " WHERE TELEFONE = '" & ID & "' "
End If


Set RSBUSCA = server.createobject("ADODB.Recordset")
RSBUSCA.Open SQLL, Conexao

If Not RSBUSCA.EOF Then

	ID = RSBUSCA("Telefone")
	NomeCliente = RSBUSCA("NomeCliente")
	DDDTelefone = RSBUSCA("Telefone")
	Valor = RSBUSCA("Valor")
	DataPedido = RSBUSCA("Data")
	Filial = RSBUSCA("Filial")		
	Loja = RSBUSCA("Loja")

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
      <p align="center"><font face="Verdana" size="3"><b>Ficha de Atendimento - Pos-Venda - <%=Loja%></b></font></td>
  </tr>
</table>
<br>

<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td colspan="2" width="50%" align="center">Nome Cliente</td>
    <td colspan="2" width="50%" align="center">DDD e Telefone</td>
  </tr>
  <tr>
    <td colspan="2" width="50%" align="center"><p align="center"><%=NomeCliente%></td>
    <td colspan="2" width="50%" align="center"><p align="center"><%=DDDTelefone%></td>
  </tr>

  <tr bgcolor="#EBF3F1">
    <td width="16%" align="center">Valor</td>
    <td width="16%" align="center">Data Pedido</td>
    <td width="16%" align="center">Filial</td>
    <td width="16%" align="center">Loja</td>
  </tr>
  <tr>
    <td width="16%" align="center"><b>R$ <%=Valor%></b></td>
    <td width="16%" align="center"><%=DataPedido%></td>
    <td width="16%" align="center"><%=Filial%></td>
    <td width="16%" align="center"><%=Loja%></td>
  </tr>


</table>
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td width="25%" align="center"><p align="center"><b>Atendimento Call Center</b></td>
    <td width="25%" align="center"><p align="center"><b>Atendimento Motoboy</b></td>
	<td width="25%" align="center"><p align="center"><b>Tempo de Entrega</b></td>
	<td width="25%" align="center"><p align="center"><b>Produto</b></td>
  </tr>
  <tr>
    <td align="left">  <p align="center">

<select name="txtCallCenter" type="text" class="textoform" size="1" >
    <option value="0" selected>Selecione aqui</option>
<option value="Otimo">Ótimo</option>
<option value="Bom">Bom</option>
<option value="Regular">Regular</option>
<option value="Ruim">Ruim</option>
<option value="Pessimo">Péssimo</option>
   </select>    
    </td>
    <td align="left">  <p align="center">
<select name="txtMotoBoy" type="text" class="textoform" size="1" >
    <option value="0" selected>Selecione aqui</option>
<option value="Otimo">Ótimo</option>
<option value="Bom">Bom</option>
<option value="Regular">Regular</option>
<option value="Ruim">Ruim</option>
<option value="Pessimo">Péssimo</option>
   </select>      
     </td>

    <td align="left">  <p align="center">

<select name="txtTempo" type="text" class="textoform" size="1" >
    <option value="0" selected>Selecione aqui</option>
<option value="Otimo">Ótimo</option>
<option value="Bom">Bom</option>
<option value="Regular">Regular</option>
<option value="Ruim">Ruim</option>
<option value="Pessimo">Péssimo</option>
   </select>    
    </td>
    <td align="left">  <p align="center">

<select name="txtProduto" type="text" class="textoform" size="1" >
    <option value="0" selected>Selecione aqui</option>
<option value="Otimo">Ótimo</option>
<option value="Bom">Bom</option>
<option value="Regular">Regular</option>
<option value="Ruim">Ruim</option>
<option value="Pessimo">Péssimo</option>
   </select>    
    </td>	

 </table>
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">

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
    <input type="hidden" name="txtDDDTelefone" value="<%=DDDTelefone%>">
	<input type="hidden" name="txtNomeCliente" value="<%=NomeCliente%>">
	<input type="hidden" name="txtFilial" value="<%=Filial%>">
	<input type="hidden" name="txtLoja" value="<%=Loja%>">
	<input type="hidden" name="txtDataPedido" value="<%=DataPedido%>">
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
If Request("txtLoja") = "Pastel na Hora" Then
	SQL4 = "INSERT INTO DNA_PastelNaHora..tbFichasAtendimento_PosVenda "
End If
	SQL4 = SQL4 & "(DDDTelefone, NomeCliente, Filial, Loja, DataPedido, DataCriacao, Observacao, CallCenter, Motoboy, Tempo, Produto, Usuario) VALUES ("& Request("txtDDDTelefone") &", '"& Request("txtNomeCliente") &"', '"& Request("txtFilial") &"', '"& Request("txtLoja") &"', '"& Request("txtDataPedido") &"', getdate(), '"& Observa &"', '"& Request("txtCallCenter") &"', '"& Request("txtMotoboy") &"', '"& Request("txtTempo") &"', '"& Request("txtProduto") &"', '"&User&"' )"
response.write SQL4
Conexao.Execute(SQL4)

	If Err.number = 0 Then
		Response.Write "<script language='javascript'>"
If Request("txtDetResultado") = "Pedido Realizado" Then
		Response.Write "alert('Pesquisa de Pos-Venda Realizada com Sucesso!');"
		Response.Write "location.href='verificacao_AtivoPosVenda.asp?Enviar=S';"
Else
		Response.Write "alert('A Pesquisa de Pos-Venda foi atualizada com sucesso.');"
		Response.Write "location.href='verificacao_AtivoPosVenda.asp?Enviar=S';"
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

If Request("txtDDDTelefone") <> "" and Enviar <> "S" Then
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

