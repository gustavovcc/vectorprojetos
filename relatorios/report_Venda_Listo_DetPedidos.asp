<%
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma", "No-Cache"

Dim Enviar, Data, VendedorCadastro, SupervisorCadastro, status, loja, Intervalo
Enviar = Request.QueryString("Enviar")
Data = Request.QueryString("Data")
Intervalo = Request.QueryString("Intervalo")
VendedorCadastro = Request.QueryString("VendedorCadastro")
SupervisorCadastro = Request.QueryString("SupervisorCadastro")
Status = Request.QueryString("status")
loja = Request.QueryString("loja")
User = Session("usuario")
If User = "" Then
	Response.Write "<script language='javascript'>"
	Response.Write "alert('Efetue o logon novamente.');"
	Response.Write "parent.parent.top.location.href='default.asp';"
	Response.Write "</script>"
	Response.End
End If
If Status = "'1'" Then StatusPed = "Aberto" End If
If Status = "'2'" Then StatusPed = "Pendente" End If
If Status = "'3'" Then StatusPed = "Producao" End If
If Status = "'4'" Then StatusPed = "Embalado" End If
If Status = "'5'" Then StatusPed = "Entrega" End If
If Status = "'6'" Then StatusPed = "Entregue" End If
If Status = "'7'" Then StatusPed = "Finalizado" End If
If Status = "'8'" Then StatusPed = "Cancelado" End If
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
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="initialize();calcRoute()">
<%
Sub Historico
%>
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td colspan="15"> <p align="center"><b>Detalhamento Pedidos do Intervalo - <%=Intervalo%> - Status: <%=StatusPed%></b></td>
  </tr>
  <tr bgcolor="#EBF3F1">
    <td width="2%" align="center"><p align="center"><i>Detalhe<br>Cliente</i></td>
    <td width="15%" align="center"><p align="center"><i>Intervalo</i></td>
    <td width="15%" align="center"><p align="center"><i>Producao</i></td>
    <td width="15%" align="center"><p align="center"><i>Data Status</i></td>
    <td width="5%" align="center"><p align="center"><i>Controle</i></td>
    <td width="10%" align="center"><p align="center"><i>Loja</i></td>
    <td width="10%" align="center"><p align="center"><i>Cliente</i></td>
    <td width="15%" align="center"><p align="center"><i>Vendedor</i></td>
      </tr>

<!--#include file="AbreConexao.asp"-->
<%
If Len(Data) < 8 Then
	Data = Split(Data,"/")
	Mes = Data(0)
	Ano = Data(1)
Else
	Data = Split(Data,"-")
	Ano = Data(0)
	Mes = Data(1)
	Dia = Data(2)

End If

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
SQL = SQL & " Status, seq_filial, nome_vendedor, data_producao, nome_loja, id, nome_cliente, dh_status "
SQL = SQL & " FROM tbPedidosListo "
SQL = SQL & " WHERE tipo_delivery in (0, 1) and "
If Len(Request.QueryString("Data")) < 8 Then
SQL = SQL & " MONTH(data_producao) = " & Mes & " AND YEAR(data_producao) = " & Ano & " "
Else
SQL = SQL & " data_producao between CONVERT(DATETIME,'"&Ano&"-"&Mes&"-"&Dia&" 00:00:00',102) and CONVERT(DATETIME,'"&Ano&"-"&Mes&"-"&Dia&" 23:59:00',102)"
End If
If Intervalo <> "Todos" Then
SQL = SQL & " and              case "
SQL = SQL & "             when cast(right(left((Convert(VarChar(8), data_producao, 108)),5),2) as Int) < 30 then left((Convert(VarChar(8), data_producao, 108)),2)  "
SQL = SQL & "             + ':00'   "
SQL = SQL & "             + ' Até '  "
SQL = SQL & "             + left((Convert(VarChar(8), data_producao, 108)),2) + ':30' "
SQL = SQL & "             when cast(right(left((Convert(VarChar(8), data_producao, 108)),5),2) as Int) >= 30 then left((Convert(VarChar(8), data_producao, 108)),2) + ':30' + ' Até '  "
SQL = SQL & "             + right('0' + cast(cast(left((Convert(VarChar(8), data_producao, 108)),2) as int) + 1 as varchar),2) "
SQL = SQL & "             + ':00' "
SQL = SQL & "             end = '"&intervalo&"' "
End If
If Loja <> "0" Then
SQL = SQL & " and nome_loja = '"&loja&"' "
End If
If Status <> "0" Then
SQL = SQL & " and status = "&status&" "
End If
SQL = SQL & " order by data_producao "

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
    <td width="2%" align="center"><p align="center"><a href="report_Venda_ContaCorrente_DetCliente.asp?User=<%=User%>&ID=<%=RSBuscas("ID")%>">
    <img src="../imagens/viewmag.png" width="16" height="16" border="0"></a></td>
</td>
    <td width="15%" align="center"><p align="center"><%=RSBuscas("Hora")%>&nbsp;</td>
    <td width="15%" align="center"><%=RSBuscas("data_producao")%></td>
    <td width="15%" align="center"><p align="center"><%=RSBuscas("dh_status")%></td>
    <td width="5%" align="center"><p align="center"><%=RSBuscas("seq_filial")%>&nbsp;</td>
    <td width="10%" align="center"><p align="center"><%=RSBuscas("nome_loja")%>&nbsp;</td>
    <td width="15%" align="center"><p align="center"><%=RSBuscas("nome_cliente")%>&nbsp;</td>
    <td width="15%" align="center"><p align="center"><%=RSBuscas("nome_vendedor")%>&nbsp;</td>
   </tr>
  <%
	RSBUSCAS.Movenext
Loop

End Sub
If Enviar <> "" Then
	Historico
End If
%>

</table>
</body>

