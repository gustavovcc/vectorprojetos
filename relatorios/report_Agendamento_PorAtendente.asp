<%
Session.LCID = 1046

Server.ScriptTimeout = 3600
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma", "No-Cache"

Dim Inicio, Retencao_Total, Cancelamento_Total, BPR_Total, Percentual_total, Cor_total, Nome_Supervisor, User, Excel, i, Resposta, Acesso, Alterada, PendenteBoleto, DiasTrabalhados, Tx_AuditoriaAgente
User = Trim(Session("usuario"))
Acesso = Session("Acesso")
Excel = Request.QueryString("Excel")
Inicio = "Consolidado"
Data = Request.QueryString("Data")
Vendedor = Request.QueryString("Vendedor")
Supervisor = Request.QueryString("Supervisor")
Gerente = Request.QueryString("Gerente")
Quartil_Anterior = 0

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
<form name="Consulta" action="report_Agendamento_PorAtendente.asp" method="get" onSubmit="return v(this)">
  <table width="100%"  border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
    <tr>
      <td width="13%" bgcolor="#EBF3F1"><strong>Data:</strong></td>
      <td><select name="Data" class="minicaption" id="Data" >
              <option value="0">Selecione uma Op&ccedil;&atilde;o</option>
              <!--#include file="AbreConexao.asp"-->
              <%
'Consolidado de outros meses
SQL_MES = "SELECT DISTINCT MONTH(convert(datetime, data, 103)) AS MES,YEAR(convert(datetime, data, 103)) AS ANO FROM NET_ATOM_Protocolo_Tabulacao_Agendamento WITH (NOLOCK) ORDER BY YEAR(convert(datetime, data, 103)),MONTH(convert(datetime, data, 103))"
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

SQL = "SELECT DISTINCT LEFT(CONVERT(CHAR(10), convert(datetime, data, 103), 120), 10) AS DATA FROM NET_ATOM_Protocolo_Tabulacao_Agendamento WITH (NOLOCK) WHERE MONTH(convert(datetime, data, 103))= "& Mes_busca &" AND YEAR(convert(datetime, data, 103)) = YEAR(GETDATE()) ORDER BY Data "
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
              <option value="0">Todos</option>
              <!--#include file="AbreConexao.asp"-->
              <%
'Consolidado de outros meses
SQL_MES = " SELECT DISTINCT Supervisor as Nome FROM SIS_USUARIOS WHERE SUPERVISOR IS NOT NULL "
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
      </select></td>
    </tr>
    <tr>
      <td bgcolor="#EBF3F1"><strong>Vendedor:</strong></td>
      <td><select name="Vendedor" class="minicaption" id="Vendedor" >
              <option value="0">Todos</option>
              <!--#include file="AbreConexao.asp"-->
              <%
'Consolidado de outros meses
SQL_MES = " SELECT DISTINCT [LOGIN OPERADOR] as Nome FROM NET_ATOM_Protocolo_Tabulacao_Agendamento "
If Acesso = "Vendas" or Acesso = "Tecnica" Then
SQL_MES = SQL_MES & " WHERE [LOGIN OPERADOR] = N'"&User&"' "
End If
SQL_MES = SQL_MES & " ORDER BY [LOGIN OPERADOR] "

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
      <p align="center"><font face="Verdana" size="3"><b>Acompanhamento Agendamento - Por Atendente</b></font>    </td>
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
SQL2 = " SELECT MAX(convert(datetime, data, 103)) AS DATA FROM NET_ATOM_Protocolo_Tabulacao_Agendamento "
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
    <td width="20%" align="left"><b><font face="Verdana" size="2" >Atendente:</font></b></td>
    <td width="80%" align="left"><b><font face="Verdana" size="2" color="#FF0000"><%=Vendedor%></font></b></td>
  </tr>
</table>
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td width="5%" align="center"><i>Rank</i></td>
    <td width="5%" align="center"><i>Quartil</i></td>
    <td width="20%" align="center"><p align="center"><i>Nome</i></td>
    <td width="10%" align="center"><div align="center">Agendamento</div></td>
    <td width="10%" align="center"><div align="center">Telefonia</div></td>
    <td width="10%" align="center"><div align="center">Atendidas</div></td>
    <td width="10%" align="center"><div align="center">Tx. Agend.</div></td>
	<td width="10%" align="center"><div align="center">Sem Sucesso</div></td>
	<td width="10%" align="center"><div align="center">Sem Possibilidade</div></td>
    <td width="10%" align="center"><div align="center">Tx. Atend.</div></td>
      </tr>
<!--#include file="AbreConexao.asp"-->
<%

SQL = " SELECT "
SQL = SQL & " Operador, "
SQL = SQL & " SUM(Agendamento) as Agendamento, "
SQL = SQL & " SUM(Atendida) as Atendidas, "
SQL = SQL & " SUM(Telefonia) as Telefonia, "
SQL = SQL & " SUM(ImprodPossivel) as ImprodPossivel, "
SQL = SQL & " SUM(ImprodImpossivel) as ImprodImpossivel, "
SQL = SQL & " SUM(Backlog) as Backlog, "
SQL = SQL & " SUM(Atendida)+ "
SQL = SQL & " SUM(Telefonia) as Total, "
SQL = SQL & " CASE (convert(decimal(20,8),ISNULL((  "
SQL = SQL & " ISNULL(SUM(Atendida),0) "
SQL = SQL & " ),0)))  "
SQL = SQL & " WHEN 0 THEN 0 ELSE ISNULL(round((convert(decimal(20,8),ISNULL(  "
SQL = SQL & " ISNULL(SUM(Agendamento),0)+ "
SQL = SQL & " ISNULL(SUM(Backlog),0) "
SQL = SQL & " ,0))) / (convert(decimal(20,8),ISNULL(  "
SQL = SQL & " ISNULL(SUM(Atendida),0) "
SQL = SQL & " ,0))) * 100,2),0) END AS TxAgendamento, "

SQL = SQL & " CASE (convert(decimal(20,8),ISNULL((  "
SQL = SQL & " ISNULL(SUM(Atendida),0) "
SQL = SQL & " ),0)))  "
SQL = SQL & " WHEN 0 THEN 0 ELSE ISNULL(round((convert(decimal(20,8),ISNULL(  "
SQL = SQL & " ISNULL(SUM(ImprodImpossivel) "
SQL = SQL & " ,0)  "
SQL = SQL & " ,0))) / (convert(decimal(20,8),ISNULL(  "
SQL = SQL & " ISNULL(SUM(Atendida),0) "
SQL = SQL & " ,0))) * 100,2),0) END AS TxImprodImpossivel, "

SQL = SQL & " CASE (convert(decimal(20,8),ISNULL((  "
SQL = SQL & " ISNULL(SUM(Atendida),0) "
SQL = SQL & " ),0)))  "
SQL = SQL & " WHEN 0 THEN 0 ELSE ISNULL(round((convert(decimal(20,8),ISNULL(  "
SQL = SQL & " ISNULL(SUM(ImprodPossivel) "
SQL = SQL & " ,0)  "
SQL = SQL & " ,0))) / (convert(decimal(20,8),ISNULL(  "
SQL = SQL & " ISNULL(SUM(Atendida),0) "
SQL = SQL & " ,0))) * 100,2),0) END AS TxImprodPossivel, "

SQL = SQL & " CASE (convert(decimal(20,8),ISNULL((  "
SQL = SQL & " ISNULL(SUM(Atendida),0)+ "
SQL = SQL & " ISNULL(SUM(Telefonia),0) "
SQL = SQL & " ),0)))  "
SQL = SQL & " WHEN 0 THEN 0 ELSE ISNULL(round((convert(decimal(20,8),ISNULL(  "
SQL = SQL & " ISNULL(SUM(Atendida) "
SQL = SQL & " ,0)  "
SQL = SQL & " ,0))) / (convert(decimal(20,8),ISNULL(  "
SQL = SQL & " ISNULL(SUM(Atendida),0)+ "
SQL = SQL & " ISNULL(SUM(Telefonia),0) "
SQL = SQL & " ,0))) * 100,2),0) END AS TxAtend "

SQL = SQL & " FROM ( "
SQL = SQL & " SELECT [Nome Operador] as Operador,  "
SQL = SQL & " SUM(cast(isnull(SRAG.Agendamento,0) as int)) as Agendamento,  "
SQL = SQL & " SUM(cast(isnull(SRAG.Atendida,0) as int)) as Atendida,  "
SQL = SQL & " SUM(cast(isnull(SRAG.Telefonia,0) as int)) as Telefonia,  "
SQL = SQL & " SUM(cast(isnull(SRAG.ImprodPossivel,0) as int)) as ImprodPossivel,  "
SQL = SQL & " SUM(cast(isnull(SRAG.ImprodImpossivel,0) as int)) as ImprodImpossivel,  "
SQL = SQL & " SUM(cast(isnull(SRAG.Backlog,0) as int)) as Backlog,  "
SQL = SQL & " COUNT(*) as Qtde  "
SQL = SQL & " FROM NET_ATOM_Protocolo_Tabulacao_Agendamento TAG "
SQL = SQL & " LEFT OUTER JOIN NET_ATOMAgendamento_SISResultado SRAG ON TAG.Resultado = SRAG.Resultado "
If Len(Request.QueryString("Data")) < 8 Then
SQL = SQL & " WHERE MONTH(convert(datetime, data, 103)) = " & Mes & " AND YEAR(convert(datetime, data, 103)) = " & Ano & " "
Else
	SQL = SQL & " Where convert(datetime, data, 103) between CONVERT(DATETIME,'"&Ano&"-"&Mes&"-"&Dia&" 00:00:00',102) and CONVERT(DATETIME,'"&Ano&"-"&Mes&"-"&Dia&" 23:59:00',102)"
End If
If Vendedor <> "0" Then
	SQL = SQL & " AND ([LOGIN OPERADOR] = N'"&Vendedor&"') "
End If
SQL = SQL & " AND [Nome Operador] NOT LIKE '%OPERADOR DISCADOR%' "
SQL = SQL & " GROUP BY [Nome Operador], SRAG.Agendamento, SRAG.Atendida, SRAG.Telefonia ) A "
SQL = SQL & " GROUP BY Operador "
SQL = SQL & " ORDER BY  "
SQL = SQL & " cast(LEFT(CASE (convert(decimal(20,8),ISNULL((  "
SQL = SQL & " ISNULL(SUM(Atendida),0) "
SQL = SQL & " ),0)))  "
SQL = SQL & " WHEN 0 THEN 0 ELSE ISNULL(round((convert(decimal(20,8),ISNULL(  "
SQL = SQL & " ISNULL(SUM(Agendamento),0)+ "
SQL = SQL & " ISNULL(SUM(Backlog),0) "
SQL = SQL & " ,0))) / (convert(decimal(20,8),ISNULL(  "
SQL = SQL & " ISNULL(SUM(Atendida),0) "
SQL = SQL & " ,0))) * 100,2),0) END,5) as float) DESC "

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

	Agendamento = RSBUSCAS("Agendamento")
	Agendamento = FormatNumber(Agendamento,0,-1,0,-2)
	Backlog = RSBUSCAS("Backlog")
	Backlog = FormatNumber(Backlog,0,-1,0,-2)
	ImprodPossivel = RSBUSCAS("ImprodPossivel")
	ImprodPossivel = FormatNumber(ImprodPossivel,0,-1,0,-2)
	ImprodImpossivel = RSBUSCAS("ImprodImpossivel")
	ImprodImpossivel = FormatNumber(ImprodImpossivel,0,-1,0,-2)
	Atendidas = RSBUSCAS("Atendidas")
	Atendidas = FormatNumber(Atendidas,0,-1,0,-2)
	Telefonia = RSBUSCAS("Telefonia")
	Telefonia = FormatNumber(Telefonia,0,-1,0,-2)
	Total = RSBUSCAS("Total")
	Total = FormatNumber(Total,0,-1,0,-2)
	TxAgendamento = RSBUSCAS("TxAgendamento")
	TxAgendamento = FormatNumber(TxAgendamento,2,-1,0,-2)
	TxImprodPossivel = RSBUSCAS("TxImprodPossivel")
	TxImprodPossivel = FormatNumber(TxImprodPossivel,2,-1,0,-2)
	TxImprodImpossivel = RSBUSCAS("TxImprodImpossivel")
	TxImprodImpossivel = FormatNumber(TxImprodImpossivel,2,-1,0,-2)
	TxAtend = RSBUSCAS("TxAtend")
	TxAtend = FormatNumber(TxAtend,2,-1,0,-2)

	Agendamento_TOTAL = Agendamento_TOTAL + CDbl(Agendamento)
	Backlog_TOTAL = Backlog_TOTAL + CDbl(Backlog)
	ImprodPossivel_TOTAL = ImprodPossivel_TOTAL + CDbl(ImprodPossivel)
	ImprodImpossivel_TOTAL = ImprodImpossivel_TOTAL + CDbl(ImprodImpossivel)
	Atendidas_TOTAL = Atendidas_TOTAL + CDbl(Atendidas)
	Telefonia_TOTAL = Telefonia_TOTAL + CDbl(Telefonia)
	Total_TOTAL = Total_TOTAL + CDbl(Total)

	If Atendidas_TOTAL > 0 Then
				TxAgendamento_Total = (CDbl(Agendamento_TOTAL)+CDbl(Backlog_TOTAL))/ (CDbl(Atendidas_TOTAL) )*100
	Else
	TxAgendamento_Total = 0
	End If
	TxAgendamento_Total = FormatNumber(TxAgendamento_Total,2,-1,0,-2)

	If Atendidas_TOTAL > 0 Then
				TxImprodPossivel_Total = (CDbl(ImprodPossivel_TOTAL))/ (CDbl(Atendidas_TOTAL) )*100
	Else
	TxImprodPossivel_Total = 0
	End If
	TxImprodPossivel_Total = FormatNumber(TxImprodPossivel_Total,2,-1,0,-2)

	If Atendidas_TOTAL > 0 Then
				TxImprodImPossivel_Total = (CDbl(ImprodImPossivel_TOTAL))/ (CDbl(Atendidas_TOTAL) )*100
	Else
	TxImprodImPossivel_Total = 0
	End If
	TxImprodImPossivel_Total = FormatNumber(TxImprodImPossivel_Total,2,-1,0,-2)
	
	If Total_TOTAL > 0 Then
				TxAtend_Total = (CDbl(Atendidas_TOTAL))/ (CDbl(Total_TOTAL) )*100
	Else
	TxAtend_Total = 0
	End If
	TxAtend_Total = FormatNumber(TxAtend_Total,2,-1,0,-2)	

If i = "1" Then Bonus = "" End If
If i = "2" Then Bonus = "" End If
If i = "3" Then Bonus = "" End If
If i = "4" Then Bonus = "" End If

If TxAgendamento > 15 Then
IndConversao = "BallGreen"
End If
If TxAgendamento = 15 Then
IndConversao = "BallYellow"
End If
If TxAgendamento < 15 Then
IndConversao = "BallRed"
End If

	If i <= Max_Quartil Then
		Quartil = "1"
	End If
	If i > Max_Quartil and i <= (Max_Quartil*2) Then
		Quartil = "2"
	End If
	If i > (Max_Quartil*2) and i <= (Max_Quartil*3) Then
		Quartil = "3"
	End If
	If i > (Max_Quartil*3) and i <= (Max_Quartil*4) Then
		Quartil = "4"
	End If

min = TMA \ 60
segundos = TMA MOD 60

min_Total = TMA_Total \ 60
segundos_Total = TMA_Total MOD 60


%>
 <tr>
    <td width="5%" align="left"><div align="center"><%=i%></div></td>
    <td width="5%" align="left"><div align="center"><%=Quartil%></div></td>    
    <td width="20%" align="left"><%=RSBuscas("Operador")%></td>
    <td width="10%" align="left"><div align="center"><%=Agendamento%></td>    
    <td width="10%" align="center"><div align="center"><%=Atendidas%></div></td>
	<td width="10%" align="center"><div align="center"><%=Telefonia%></div></td>
    <td width="10%" align="center"><div align="center"><%=TxAgendamento%>% <img src="../imagens/<%=IndConversao%>.gif" width="15" height="18" border="0"></div></td>
	<td width="10%" align="center"><div align="center"><%=TxImprodPossivel%>%</div></td>
	<td width="10%" align="center"><div align="center"><%=TxImprodImpossivel%>%</div></td>
    <td width="10%" align="center"><div align="center"><%=TxAtend%>%</div></td>
  </tr>
  <%
	RSBUSCAS.Movenext
Loop
%>
<tr bgcolor="#cdd5da">
    <td colspan="3" align="left"><div align="left"><b>Total</b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=Agendamento_Total%></b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=Atendidas_Total%></b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=Telefonia_Total%></b></div></td>
	<td width="10%" align="center"><div align="center"><b><%=TxAgendamento_Total%>%</b></div></td>
	<td width="10%" align="center"><div align="center"><b><%=TxImprodPossivel_Total%>%</b></div></td>
	<td width="10%" align="center"><div align="center"><b><%=TxImprodImpossivel_Total%>%</b></div></td>
	<td width="10%" align="center"><div align="center"><b><%=TxAtend_Total%>%</b></div></td>
 
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

