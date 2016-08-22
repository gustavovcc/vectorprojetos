<%
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma", "No-Cache"

Dim Enviar, VendedorCadastro, DDDTelefone, StatusPed
Enviar = Request.QueryString("Enviar")
DDDTelefone = Right(Request.QueryString("DDDTelefone"),8)

User = Session("usuario")
If User = "" Then
	Response.Write "<script language='javascript'>"
	Response.Write "alert('Efetue o logon novamente.');"
	Response.Write "parent.parent.top.location.href='default.asp';"
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
<title>BS Digital</title>
<link rel="stylesheet" href="../include/pgo.css" type="text/css">
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<%
Sub Historico
%>
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td colspan="15"> <p align="center"><b>Detalhamento Pedidos do Telefone - <%=DDDTelefone%></b></td>
  </tr>
  <tr bgcolor="#EBF3F1">
    <td width="10%" align="center"><p align="center"><i>Data</i></td>
    <td width="10%" align="center"><p align="center"><i>Data Status</i></td>
    <td width="5%" align="center"><p align="center"><i>Status</i></td>
    <td width="5%" align="center"><p align="center"><i>Controle</i></td>
    <td width="8%" align="center"><p align="center"><i>Loja</i></td>
    <td width="15%" align="center"><p align="center"><i>Cliente</i></td>
    <td width="15%" align="center"><p align="center"><i>Vendedor</i></td>
    <td width="5%" align="center"><p align="center"><i>Valor Conta</i></td>
      </tr>

<!--#include file="AbreConexao.asp"-->
<%

SQL = " SELECT "
SQL = SQL & "             case "
SQL = SQL & "             when cast(right(left((Convert(VarChar(8), data_producao, 108)),5),2) as Int) < 30 then left((Convert(VarChar(8), data_producao, 108)),2)  "
SQL = SQL & "             + ':00'   "
SQL = SQL & "             + ' Até '  "
SQL = SQL & "             + left((Convert(VarChar(8), data_producao, 108)),2) + ':30' "
SQL = SQL & "             when cast(right(left((Convert(VarChar(8), data_producao, 108)),5),2) as Int) >= 30 then left((Convert(VarChar(8), data_producao, 108)),2) + ':30' + ' Até '  "
SQL = SQL & "             + right('0' + cast(cast(left((Convert(VarChar(8), data_producao, 108)),2) as int) + 1 as varchar),2) "
SQL = SQL & "             + ':00' "
SQL = SQL & "             end as Hora,  "
SQL = SQL & " Status, seq_filial, nome_vendedor, data_producao, nome_loja, id, nome_cliente, dh_status, Valor_Conta "
SQL = SQL & " FROM tbPedidosListo "

SQL = SQL & " WHERE  "
SQL = SQL & " (RIGHT(left(TEL_RES,3),2)+ "
SQL = SQL & " CASE len(TEL_RES) WHEN 13 THEN left(RIGHT(TEL_RES,9),4) ELSE left(RIGHT(TEL_RES,10),5) END + "
SQL = SQL & " RIGHT(TEL_RES,4) LIKE '%"&DDDTelefone&"%'  "
SQL = SQL & " OR  "
SQL = SQL & " RIGHT(left(TEL_COM,3),2)+ "
SQL = SQL & " CASE len(TEL_COM) WHEN 13 THEN left(RIGHT(TEL_COM,9),4) ELSE left(RIGHT(TEL_COM,10),5) END + "
SQL = SQL & " RIGHT(TEL_COM,4) LIKE '%"&DDDTelefone&"%' "
SQL = SQL & " OR  "
SQL = SQL & " RIGHT(left(TEL_CEL,3),2)+ "
SQL = SQL & " CASE len(TEL_CEL) WHEN 13 THEN left(RIGHT(TEL_CEL,9),4) ELSE left(RIGHT(TEL_CEL,10),5) END + "
SQL = SQL & " RIGHT(TEL_CEL,4) LIKE '%"&DDDTelefone&"%') "
SQL = SQL & " order by data_producao ASC "

Set RSBUSCAS = server.createobject("ADODB.Recordset")
'RESPONSE.WRITE SQL
RSBUSCAS.Open SQL, Conexao
	i = 0
Do While Not RSBUSCAS.EOF
	i = i + 1

If RSBuscas("Status") = 1 Then StatusPed = "Aberto" End If
If RSBuscas("Status") = 2 Then StatusPed = "Pendente" End If
If RSBuscas("Status") = 3 Then StatusPed = "Producao" End If
If RSBuscas("Status") = 4 Then StatusPed = "Embalado" End If
If RSBuscas("Status") = 5 Then StatusPed = "Entrega" End If
If RSBuscas("Status") = 6 Then StatusPed = "Entregue" End If
If RSBuscas("Status") = 7 Then StatusPed = "Finalizado" End If
If RSBuscas("Status") = 8 Then StatusPed = "Cancelado" End If

%>
  <tr>
    <td width="10%" align="center"><%=RSBuscas("data_producao")%></td>
    <td width="10%" align="center"><p align="center"><%=RSBuscas("dh_status")%></td>
    <td width="5%" align="center"><p align="center"><%=StatusPed%>&nbsp;</td>
    <td width="5%" align="center"><p align="center"><%=RSBuscas("seq_filial")%>&nbsp;</td>
    <td width="8%" align="center"><p align="center"><%=RSBuscas("nome_loja")%>&nbsp;</td>
    <td width="15%" align="center"><p align="center"><%=RSBuscas("nome_cliente")%>&nbsp;</td>
    <td width="15%" align="center"><p align="center"><%=RSBuscas("nome_vendedor")%>&nbsp;</td>
    <td width="5%" align="center"><p align="center">R$ <%=RSBuscas("Valor_Conta")%>&nbsp;</td>
   </tr>
  <%
	RSBUSCAS.Movenext
Loop

End Sub
If Enviar = "" Then
	Historico
End If
%>

</table>
</body>

