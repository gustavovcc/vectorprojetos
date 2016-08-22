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

Dim Obs, Atualiza, Enviar, Observa, Data, ID, NomeFantasia, RazaoSocial, CNPJ, Bairro, Cidade, CEP, Contato, Telefone, Vendedor, Revenda, Consultor
Dim DataCadastro, DataPrimeiraVisita, DataUltimaVisita, UltimoFunil, UltimoDetFunil, ProximaVisita, ProdutoCliente, Operadora, Tempo

ID = request.querystring("ID")
IDVisita = request.querystring("IDVisita")
Enviar = request.querystring("Enviar")
Atualiza = request.querystring("Atualiza")
%>
<html>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>WeDo</title>
<link rel="stylesheet" href="../include/pgo.css" type="text/css">
</head>
<form method="POST" name="frmDados">
<%
Sub HistoricoAtendimentos
%>
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td colspan="10"> <p align="center"><b>Histórico Atendimentos</b></td>
  </tr>
  <tr bgcolor="#EBF3F1">
    <td width="5%" align="center"> <p align="center"><i>Data Atendimento</i></td>
    <td width="5%" align="center"> <p align="center"><i>Resultado Chamada</i></td>
    <td width="10%" align="center"><p align="center"><i>Detatlhe Resultado</i></td>
    <td width="10%" align="center"><p align="center"><i>Atendente</i></td>
    <td width="10%" align="center"><p align="center"><i>Observação Atendimento</i></td>
     </tr>

<!--#include file="AbreConexao.asp"-->
<%

SQL = "SELECT * FROM tbFichasAtendimento "
SQL = SQL & " where IDProspects = "& ID &" "
SQL = SQL & "ORDER BY DataCriacao ASC "
Set RSBUSCAS = server.createobject("ADODB.Recordset")
'RESPONSE.WRITE SQL
RSBUSCAS.Open SQL, Conexao

Do While Not RSBUSCAS.EOF
%>
  <tr>
    <td width="5%" align="center"><p align="center"><%=RSBuscas("DataCriacao")%>&nbsp;</td>
    <td width="5%" align="center"><p align="center"><%=RSBuscas("ResultadoChamada")%>&nbsp;</td>
    <td width="10%" align="center"><p align="center"><%=RSBuscas("DetResultado")%>&nbsp;</td>
    <td width="10%" align="center"><p align="center"><%=RSBuscas("ResponsavelCriacao")%>&nbsp;</td>
    <td width="10%" align="center"><p align="center"><%=RSBuscas("Observacao")%>&nbsp;</td>
  </tr>
  <%
	RSBUSCAS.Movenext
Loop

End Sub
If ID <> "" and Enviar <> "S" Then
	HistoricoAtendimentos
End If
%>

</table>

</body>

