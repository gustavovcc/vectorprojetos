<%
Session.LCID = 1046

Server.ScriptTimeout = 3600
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma", "No-Cache"

Dim Inicio, Retencao_Total, Cancelamento_Total, BPR_Total, Percentual_total, Cor_total, Nome_Supervisor, User, Excel, i, Resposta, Acesso, Alterada, PendenteBoleto, DiasTrabalhados, Tx_AuditoriaAgente, Loja
User = Trim(Session("usuario"))
Acesso = Session("Acesso")
Excel = Request.QueryString("Excel")
Inicio = "Consolidado"
Data = Request.QueryString("Data")
Loja = Request.QueryString("Loja")
Vendedor = Request.QueryString("Vendedor")
Supervisor = Request.QueryString("Supervisor")

If User = "" Then
	Response.Write "<script language='javascript'>"
	Response.Write "alert('Efetue o logon novamente.');"
	Response.Write "parent.parent.top.location.href='../default.asp';"
	Response.Write "</script>"
	Response.End
End If


If Excel = "S" Then
	Response.ContentType = "Application/vnd.ms-excel"
End If
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="refresh" content="1800">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Ranking Mensal</title>
<link rel="stylesheet" href="../include/pgo.css" type="text/css">
<script>
function v(frm){
	f = frm;
	if(f.Data.value=='0'){
	alert('Escolha uma data para geração do relatório.');
	f.Data.focus();
	return false;
	}
	if(f.Segmento.value=='0'){
		alert('Escolha um segmento para geração do relatório.');
		f.Segmento.focus();
		return false;
	}
	mostra_destaque(emdestaque);
	return true;
}
function esconde_destaque( menu ) {
menu.style.visibility="hidden"
}
function mostra_destaque( menu ) {
menu.style.visibility="visible"
}
function Atualizar()
{
    if (NewWindow != null) {
       NewWindow.close();
       NewWindow = null;
       NewWindow = window.open("altera.asp","JanCalendario","Height=230,width=250,top=200,left=300");
       }
    else{
       NewWindow = window.open("altera.asp","JanCalendario","Height=230,width=250,top=200,left=300");
       }
}
</script>
</head>
<body>
<form name="Consulta" action="report_Venda_Listo.asp" method="get" onSubmit="return v(this)">
  <table width="100%"  border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
    <tr>
      <td width="13%" bgcolor="#EBF3F1"><strong>Data:</strong></td>
      <td><select name="Data" class="minicaption" id="Data" >
              <option value="0">Selecione uma Op&ccedil;&atilde;o</option>
              <!--#include file="AbreConexao.asp"-->
              <%
'Consolidado de outros meses
SQL_MES = "SELECT DISTINCT MONTH(data_producao) AS MES,YEAR(data_producao) AS ANO FROM tbPedidosListo WITH (NOLOCK) where data_Producao is not null ORDER BY YEAR(data_producao),MONTH(data_producao)"
Set RSBUSCA = server.createobject("ADODB.Recordset")
RSBUSCA.Open SQL_MES, Conexao

Do While Not RSBUSCA.EOF
	Nome_mes_busca = MonthName(RSBUSCA("MES"))
%>
              <option value="<%=RSBUSCA("MES")%>/<%=RSBUSCA("ANO")%>">Consolidado de <%=Nome_mes_busca%>/<%=RSBUSCA("ANO")%></option>
              <%
	RSBUSCA.MoveNext
Loop

If Day(Now) = 1 Then
	Mes_busca = (Month(Now)-1)
Else
	Mes_busca = Month(Now)
End If

SQL = "SELECT DISTINCT LEFT(CONVERT(CHAR(10), data_Producao, 120), 10) AS DATA FROM tbPedidosListo WITH (NOLOCK) WHERE MONTH(data_Producao)= "& Mes_busca &" AND YEAR(data_Producao) = YEAR(GETDATE()) ORDER BY Data "
response.write sql
Set RSBUSCA = server.createobject("ADODB.Recordset")
RSBUSCA.Open SQL, Conexao

Do While Not RSBUSCA.EOF
%>
              <option value="<%=RSBUSCA("DATA")%>"><%=RSBUSCA("DATA")%></option>
              <%
	RSBUSCA.MoveNext
Loop
	RSBUSCA.Close
	Set SQL = Nothing
	Set Conexao = Nothing
%>
      </select></td>
    </tr>

    <tr>
      <td bgcolor="#EBF3F1"><strong>Supervisor:</strong></td>
      <td><select name="Supervisor" class="minicaption" id="Tecnico" >
<% If Acesso = "Administrador" or Acesso = "Vendas" or Acesso = "Auditor" or Acesso = "Back-Office" Then %>
              <option value="0">Todos</option>
<%End If%>
<% If Acesso <> "Vendas" Then %>
              <!--#include file="AbreConexao.asp"-->
              <%
'Consolidado de outros meses
SQL_MES = " SELECT DISTINCT Supervisor as Nome FROM DNA_PizzaHut..SIS_USUARIOS WHERE SUPERVISOR IS NOT NULL "
If Acesso = "Supervisor" Then
SQL_MES = SQL_MES & " and Supervisor  = N'"&User&"' "
End If
SQL_MES = SQL_MES & " ORDER BY SUPERVISOR "

Set RSBUSCA = server.createobject("ADODB.Recordset")
RSBUSCA.Open SQL_MES, Conexao

Do While Not RSBUSCA.EOF

%>
              <option value="<%=RSBUSCA("Nome")%>"><%=RSBUSCA("Nome")%></option>
              <%
	RSBUSCA.MoveNext
Loop

	Set RSBUSCA = Nothing
	Set Conexao = Nothing
%>
<%End If%>
      </select></td>
    </tr>
    <tr>
      <td bgcolor="#EBF3F1"><strong>Vendedor:</strong></td>
      <td><select name="Vendedor" class="minicaption" id="Vendedor" >
<% If Acesso = "Administrador" or Acesso = "Supervisor" or Acesso = "Auditor" or Acesso = "Back-Office" Then %>
              <option value="0">Todos</option>
<%End If%>
              <!--#include file="AbreConexao.asp"-->
              <%
'Consolidado de outros meses
SQL_MES = " SELECT DISTINCT ResponsavelCriacao as Nome FROM tbFichasAtendimento "
If Acesso = "Vendas" or Acesso = "Tecnica" Then
SQL_MES = SQL_MES & " WHERE ResponsavelCriacao = N'"&User&"' "
End If
If Acesso = "Supervisor"  Then
SQL_MES = SQL_MES & " Where ResponsavelCriacao in (Select usuario from dna_PizzaHut..SIS_Usuarios where Supervisor = N'"&User&"') "
End If
SQL_MES = SQL_MES & " ORDER BY Nome "

Set RSBUSCA = server.createobject("ADODB.Recordset")
RSBUSCA.Open SQL_MES, Conexao

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
      <td bgcolor="#EBF3F1"><strong>Loja:</strong></td>
      <td><select name="Loja" class="minicaption" id="Loja" >
              <option value="0">Todos</option>
              <!--#include file="AbreConexao.asp"-->
              <%
'Consolidado de outros meses
SQL_MES = " SELECT DISTINCT nome_loja as Nome FROM tbPedidosListo "
SQL_MES = SQL_MES & " ORDER BY Nome "

Set RSBUSCA = server.createobject("ADODB.Recordset")
RSBUSCA.Open SQL_MES, Conexao

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

<%
'//Acesso Administrador
'Acesso = Session("Acesso")
'If Acesso = "Administrador" Then
%>
     <tr>
      <td>&nbsp;</td>
      <td><input name="Submit" type="submit" class="formbutton" value="Buscar"></td>
    </tr>

  </table>
</form>
<%Sub Busca_Retencao
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

%>
<table border="0" width="100%">
  <tr>
    <td width="84%">
      <p align="center"><font face="Verdana" size="3"><b>Acompanhamento Vendas</b></font>    </td>
    <td width="16%">
	<!--
      <p align="center"><b><font color="#FF0000"><img border="0" src="imagens/excel.jpg" width="16" height="16">
      <a target="blank" class="topo" href="ranking_result_novo.asp?Data=<%=Request.QueryString("Data")%>&supervisor=<%=Request.QueryString("supervisor")%>&Excel=S">Exportar p/ Excel</a></font></b>
	  -->
	</td>
  </tr>
</table>

<table border="0" width="100%">
  <tr>
    <td colspan="2"><font face="Verdana" size="2"><b>Data:</b></font><font color="red" face="Verdana" size="2"><b>
      <%If Len(Request.QueryString("Data")) < 8  Then Response.Write("Consolidado do mês de " & MonthName(Mes) &"/"& Ano) Else Response.Write(Dia&"/"&Mes&"/"&Ano) End If%>
    </b></font><font color="red" face="Verdana" size="2">&nbsp;</font></td>
  </tr>
  <!--#include file="AbreConexao.asp"-->
  <%
SQL2 = " SELECT MAX(dh_status) AS DATA FROM tbPedidosListo "
	Set RSMAXIMO = server.createobject("ADODB.Recordset")
	RSMAXIMO.Open SQL2, Conexao
	If Not RSMAXIMO.EOF Then
		Maximo = Trim(RSMAXIMO("DATA"))
	End If

Set RSMAXIMO = Nothing
Set Conexao = Nothing


%>
  <tr>
    <td width="20%" align="left"><b><font face="Verdana" size="2" >Indicador atualizado até:</font></b></td>
    <td width="80%" align="left"><b><font face="Verdana" size="2" color="#FF0000"><%=Maximo%></font></b></td>
  </tr>
  <tr>
    <td width="20%" align="left"><b><font face="Verdana" size="2" >Vendedor:</font></b></td>
    <td width="80%" align="left"><b><font face="Verdana" size="2" color="#FF0000"><%=Vendedor%></font></b></td>
  </tr>
  <tr>
    <td width="20%" align="left"><b><font face="Verdana" size="2" >Loja:</font></b></td>
    <td width="80%" align="left"><b><font face="Verdana" size="2" color="#FF0000"><%=Loja%></font></b></td>
  </tr>

</table>
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td width="5%" align="center"><i>Intervalo</i></td>
    <td width="5%" align="center"><div align="center">Aberto</div></td>
    <td width="5%" align="center"><div align="center">Pendente</div></td>
    <td width="5%" align="center"><div align="center">Producao</div></td>
    <td width="5%" align="center"><div align="center">Embalado</div></td>
    <td width="5%" align="center"><div align="center">Entrega</div></td>
    <td width="5%" align="center"><div align="center">Entregue</div></td>
    <td width="5%" align="center"><div align="center">Finalizado</div></td>
    <td width="5%" align="center"><div align="center">Cancelado</div></td>
    <td width="5%" align="center"><div align="center">Finalizaveis</div></td>
      </tr>
<!--#include file="AbreConexao.asp"-->
<%


SQL = " SELECT "
SQL = SQL & "             case  "
SQL = SQL & "             when cast(right(left((Convert(VarChar(8), data_producao, 108)),5),2) as Int) < 30 then left((Convert(VarChar(8), data_producao, 108)),2)  "
SQL = SQL & "             + ':00'   "
SQL = SQL & "             + ' Até '  "
SQL = SQL & "             + left((Convert(VarChar(8), data_producao, 108)),2) + ':30' "
SQL = SQL & "             when cast(right(left((Convert(VarChar(8), data_producao, 108)),5),2) as Int) >= 30 then left((Convert(VarChar(8), data_producao, 108)),2) + ':30' + ' Até '  "
SQL = SQL & "             + right('0' + cast(cast(left((Convert(VarChar(8), data_producao, 108)),2) as int) + 1 as varchar),2) "
SQL = SQL & "             + ':00' "
SQL = SQL & "             end as Hora, "
SQL = SQL & " SUM(CASE WHEN status = '1' THEN 1 ELSE 0 END) as Aberto, "
SQL = SQL & " SUM(CASE WHEN status = '2' THEN 1 ELSE 0 END) as Pendente, "
SQL = SQL & " SUM(CASE WHEN status = '3' THEN 1 ELSE 0 END) as Producao, "
SQL = SQL & " SUM(CASE WHEN status = '4' THEN 1 ELSE 0 END) as Embalado, "
SQL = SQL & " SUM(CASE WHEN status = '5' THEN 1 ELSE 0 END) as Entrega, "
SQL = SQL & " SUM(CASE WHEN status = '6' THEN 1 ELSE 0 END) as Entregue, "
SQL = SQL & " SUM(CASE WHEN status = '7' THEN 1 ELSE 0 END) as Finalizado, "
SQL = SQL & " SUM(CASE WHEN status = '8' THEN 1 ELSE 0 END) as Cancelado "
SQL = SQL & " FROM tbPedidosListo "

If Len(Request.QueryString("Data")) < 8 Then
SQL = SQL & " WHERE tipo_delivery in (0, 1) and MONTH(Data_producao) = " & Mes & " AND YEAR(Data_producao) = " & Ano & " "
Else
	SQL = SQL & " Where tipo_delivery in (0, 1) and Data_producao between CONVERT(DATETIME,'"&Ano&"-"&Mes&"-"&Dia&" 00:00:00',102) and CONVERT(DATETIME,'"&Ano&"-"&Mes&"-"&Dia&" 23:59:00',102)"
End If
If Vendedor <> "0" Then
	SQL = SQL & " AND (usuario = N'"&Vendedor&"') "
End If
If Loja <> "0" Then
	SQL = SQL & " AND (nome_loja = N'"&Loja&"') "
End If
	SQL = SQL & " AND (ID_USUARIO_VENDEDOR not in ('admin', 'myrlla', 'marina.silva', 'ruslana.pires')) "

SQL = SQL & " GROUP BY  "
SQL = SQL & " case  "
SQL = SQL & " when cast(right(left((Convert(VarChar(8), data_producao, 108)),5),2) as Int) < 30 then left((Convert(VarChar(8), data_producao, 108)),2)  "
SQL = SQL & " + ':00'   "
SQL = SQL & " + ' Até '  "
SQL = SQL & " + left((Convert(VarChar(8), data_producao, 108)),2) + ':30' "
SQL = SQL & " when cast(right(left((Convert(VarChar(8), data_producao, 108)),5),2) as Int) >= 30 then left((Convert(VarChar(8), data_producao, 108)),2) + ':30' + ' Até '  "
SQL = SQL & " + right('0' + cast(cast(left((Convert(VarChar(8), data_producao, 108)),2) as int) + 1 as varchar),2) "
SQL = SQL & " + ':00' "
SQL = SQL & " end "


'response.write sql
Set RSBUSCAS = server.createobject("ADODB.Recordset")
RSBUSCAS.CursorType = 0
RSBUSCAS.CursorLocation = 3
RSBUSCAS.Open SQL, Conexao
Quantidade = RSBUSCAS.RecordCount

If Quantidade > 0 Then
	Max_Quartil = (Quantidade/4)
End If


	i = 0
Do While Not RSBUSCAS.EOF
	i = i + 1

	Aberto = RSBUSCAS("Aberto")
	Aberto = FormatNumber(Aberto,0,-1,0,-2)
	Pendente = RSBUSCAS("Pendente")
	Pendente = FormatNumber(Pendente,0,-1,0,-2)
	Producao = RSBUSCAS("Producao")
	Producao = FormatNumber(Producao,0,-1,0,-2)
	Embalado = RSBUSCAS("Embalado")
	Embalado = FormatNumber(Embalado,0,-1,0,-2)
	Entrega = RSBUSCAS("Entrega")
	Entrega = FormatNumber(Entrega ,0,-1,0,-2)
	Entregue = RSBUSCAS("Entregue")
	Entregue = FormatNumber(Entregue,3,-1,0,-2)
	Finalizado = RSBUSCAS("Finalizado")
	Finalizado = FormatNumber(Finalizado,3,-1,0,-2)
	Cancelado = RSBUSCAS("Cancelado")
	Cancelado = FormatNumber(Cancelado,3,-1,0,-2)

	Finalizaveis = RSBUSCAS("Producao")+RSBUSCAS("Embalado")+RSBUSCAS("Entrega")+RSBUSCAS("Entregue")+RSBUSCAS("Finalizado")
	Finalizaveis = FormatNumber(Finalizaveis,0,-1,0,-2)

	Aberto_TOTAL = Aberto_TOTAL + CDbl(Aberto)
	Pendente_TOTAL = Pendente_TOTAL + CDbl(Pendente)
	Producao_TOTAL = Producao_TOTAL + CDbl(Producao)
	Embalado_TOTAL = Embalado_TOTAL + CDbl(Embalado)
	Entrega_TOTAL = Entrega_TOTAL + CDbl(Entrega)
	Entregue_TOTAL = Entregue_TOTAL + CDbl(Entregue)
	Finalizado_TOTAL = Finalizado_TOTAL + CDbl(Finalizado)
	Cancelado_TOTAL = Cancelado_TOTAL + CDbl(Cancelado)
	Finalizaveis_TOTAL = Finalizaveis_TOTAL + CDbl(Finalizaveis)


%>
 <tr>
    <td width="5%" align="left"><%=RSBuscas("hora")%></td>
    <td width="5%" align="left"><div align="center"><a class="topo" href="report_Venda_Listo_DetPedidos.asp?Intervalo=<%=RSBUSCAS("hora")%>&Loja=<%=loja%>&Status='1'&Data=<%=Request.QueryString("Data")%>&Enviar=S"><%=RSBuscas("Aberto")%></a></td>    
    <td width="5%" align="center"><div align="center"><a class="topo" href="report_Venda_Listo_DetPedidos.asp?Intervalo=<%=RSBUSCAS("hora")%>&Loja=<%=loja%>&Status='2'&Data=<%=Request.QueryString("Data")%>&Enviar=S"><%=RSBuscas("Pendente")%></a></div></td>
    <td width="5%" align="center"><div align="center"><a class="topo" href="report_Venda_Listo_DetPedidos.asp?Intervalo=<%=RSBUSCAS("hora")%>&Loja=<%=loja%>&Status='3'&Data=<%=Request.QueryString("Data")%>&Enviar=S"><%=RSBuscas("Producao")%></a></div></td>
    <td width="5%" align="center"><div align="center"><a class="topo" href="report_Venda_Listo_DetPedidos.asp?Intervalo=<%=RSBUSCAS("hora")%>&Loja=<%=loja%>&Status='4'&Data=<%=Request.QueryString("Data")%>&Enviar=S"><%=RSBuscas("Embalado")%></a></div></td>
    <td width="5%" align="center"><div align="center"><a class="topo" href="report_Venda_Listo_DetPedidos.asp?Intervalo=<%=RSBUSCAS("hora")%>&Loja=<%=loja%>&Status='5'&Data=<%=Request.QueryString("Data")%>&Enviar=S"><%=RSBuscas("Entrega")%></a></div></td>
    <td width="5%" align="center"><div align="center"><a class="topo" href="report_Venda_Listo_DetPedidos.asp?Intervalo=<%=RSBUSCAS("hora")%>&Loja=<%=loja%>&Status='6'&Data=<%=Request.QueryString("Data")%>&Enviar=S"><%=RSBuscas("Entregue")%></a></div></td>
    <td width="5%" align="center"><div align="center"><a class="topo" href="report_Venda_Listo_DetPedidos.asp?Intervalo=<%=RSBUSCAS("hora")%>&Loja=<%=loja%>&Status='7'&Data=<%=Request.QueryString("Data")%>&Enviar=S"><%=RSBuscas("Finalizado")%></a></div></td>
    <td width="5%" align="center"><div align="center"><a class="topo" href="report_Venda_Listo_DetPedidos.asp?Intervalo=<%=RSBUSCAS("hora")%>&Loja=<%=loja%>&Status='8'&Data=<%=Request.QueryString("Data")%>&Enviar=S"><%=RSBuscas("Cancelado")%></a></div></td>
    <td width="5%" align="center"><div align="center"><%=Finalizaveis%></div></td>
  </tr>
  <%
	RSBUSCAS.Movenext
Loop
%>
<tr bgcolor="#cdd5da">
    <td colspan="1" align="left"><div align="left"><b>Total</b></div></td>
    <td width="5%" align="center"><div align="center"><b><a class="topo" href="report_Venda_Listo_DetPedidos.asp?Intervalo=Todos&Loja=<%=loja%>&Status='1'&Data=<%=Request.QueryString("Data")%>&Enviar=S"><%=Aberto_Total%></a></b></div></td>
    <td width="5%" align="center"><div align="center"><b><a class="topo" href="report_Venda_Listo_DetPedidos.asp?Intervalo=Todos&Loja=<%=loja%>&Status='2'&Data=<%=Request.QueryString("Data")%>&Enviar=S"><%=Pendente_Total%></a></b></div></td>
    <td width="5%" align="center"><div align="center"><b><a class="topo" href="report_Venda_Listo_DetPedidos.asp?Intervalo=Todos&Loja=<%=loja%>&Status='3'&Data=<%=Request.QueryString("Data")%>&Enviar=S"><%=Producao_Total%></a></b></div></td>
    <td width="5%" align="center"><div align="center"><b><a class="topo" href="report_Venda_Listo_DetPedidos.asp?Intervalo=Todos&Loja=<%=loja%>&Status='4'&Data=<%=Request.QueryString("Data")%>&Enviar=S"><%=Embalado_Total%></a></b></div></td>
    <td width="5%" align="center"><div align="center"><b><a class="topo" href="report_Venda_Listo_DetPedidos.asp?Intervalo=Todos&Loja=<%=loja%>&Status='5'&Data=<%=Request.QueryString("Data")%>&Enviar=S"><%=Entrega_Total%></a></b></div></td>
    <td width="5%" align="center"><div align="center"><b><a class="topo" href="report_Venda_Listo_DetPedidos.asp?Intervalo=Todos&Loja=<%=loja%>&Status='6'&Data=<%=Request.QueryString("Data")%>&Enviar=S"><%=Entregue_Total%></a></b></div></td>
    <td width="5%" align="center"><div align="center"><b><a class="topo" href="report_Venda_Listo_DetPedidos.asp?Intervalo=Todos&Loja=<%=loja%>&Status='7'&Data=<%=Request.QueryString("Data")%>&Enviar=S"><%=Finalizado_Total%></a></b></div></td>
    <td width="5%" align="center"><div align="center"><b><a class="topo" href="report_Venda_Listo_DetPedidos.asp?Intervalo=Todos&Loja=<%=loja%>&Status='8'&Data=<%=Request.QueryString("Data")%>&Enviar=S"><%=Cancelado_Total%></a></b></div></td>
    <td width="5%" align="center"><div align="center"><b><%=Finalizaveis_Total%></b></div></td>
  </tr>

<%
End Sub

If Request.QueryString("Data") <> "" Then
	Busca_Retencao
End If
%>
</table>
<DIV id=emdestaque style="left:350; width=300; position: absolute; top: 110; visibility: hidden; z-index: 1; border: 1px none #000000" ondrag="MM_showHideLayers('emdestaque','','show')">
<br><br><br><center>
<img src="../../imagens/carregando.gif" border="0">
</center>
</div>
</body>
</html>

