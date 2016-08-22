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
<form name="Consulta" action="report_Venda_Comissao.asp" method="get" onSubmit="return v(this)">
  <table width="100%"  border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
    <tr>
      <td width="13%" bgcolor="#EBF3F1"><strong>Data:</strong></td>
      <td><select name="Data" class="minicaption" id="Data" >
              <option value="0">Selecione uma Op&ccedil;&atilde;o</option>
              <!--#include file="AbreConexao.asp"-->
              <%
'Consolidado de outros meses
SQL_MES = "SELECT DISTINCT MONTH(DataCriacao) AS MES,YEAR(DataCriacao) AS ANO FROM tbFichasAtendimento WITH (NOLOCK) where DataCriacao is not null ORDER BY YEAR(DataCriacao),MONTH(DataCriacao)"
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

SQL = "SELECT DISTINCT LEFT(CONVERT(CHAR(10), DataCriacao, 120), 10) AS DATA FROM tbFichasAtendimento WITH (NOLOCK) WHERE MONTH(DataCriacao)= "& Mes_busca &" AND YEAR(DataCriacao) = YEAR(GETDATE()) ORDER BY Data "
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
SQL2 = " SELECT MAX(DataCriacao) AS DATA FROM tbFichasAtendimento "
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
</table>
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td width="5%" align="center"><i>Rank</i></td>
    <td width="5%" align="center"><i>Quartil</i></td>
    <td width="25%" align="center"><p align="center"><i>Nome</i></td>
    <td width="10%" align="center"><div align="center">% Marcação DNA</div></td>
    <td width="10%" align="center"><div align="center">Pedidos / Ligações</div></td>
    <td width="10%" align="center"><div align="center">Tx. Conversão</div></td>
    <td width="10%" align="center"><div align="center">TMA(s)</div></td>
    <td width="10%" align="center"><div align="center">Nota</div></td>
<%If Acesso <> "Vendas" Then %>
    <td width="10%" align="center"><div align="center">Comissão</div></td>    
<% End If %>    
      </tr>
<!--#include file="AbreConexao.asp"-->
<%

SQL = " SELECT Usuario, "
SQL = SQL & " SUM(CASE WHEN Resultado = 'Marcacoes' THEN Qtde ELSE 0 END) as Marcacoes, "
SQL = SQL & " SUM(CASE WHEN Resultado = 'Pedidos' THEN Qtde ELSE 0 END) as Pedido, "
SQL = SQL & " SUM(CASE WHEN Resultado = 'Chamadas' THEN Qtde ELSE 0 END) as Chamadas, "
SQL = SQL & " SUM(CASE WHEN Resultado = 'Tempo' THEN Qtde ELSE 0 END) as Tempo, "
SQL = SQL & " CASE SUM(CASE WHEN Resultado = 'Chamadas' THEN Qtde ELSE 0 END) "
SQL = SQL & " WHEN 0 THEN 0 ELSE "
SQL = SQL & " SUM(CASE WHEN Resultado = 'TEMPO' THEN Qtde ELSE 0 END) / "
SQL = SQL & " SUM(CASE WHEN Resultado = 'Chamadas' THEN Qtde ELSE 0 END) "
SQL = SQL & " END AS TMA, "
SQL = SQL & " CASE (convert(decimal(20,8),ISNULL(( "
SQL = SQL & " SUM(CASE WHEN Resultado = 'Chamadas' THEN Qtde ELSE 0 END) "
SQL = SQL & "  ),0))) WHEN 0 THEN 0 ELSE ISNULL(round((convert(decimal(20,8),ISNULL( "
SQL = SQL & " SUM(CASE WHEN Resultado = 'Pedidos' THEN Qtde ELSE 0 END) "
SQL = SQL & " ,0))) / (convert(decimal(20,8),ISNULL( "
SQL = SQL & " SUM(CASE WHEN Resultado = 'Chamadas' THEN Qtde ELSE 0 END) "
SQL = SQL & " ,0))) * 100,2),0) END AS Tx_Conversao, "
SQL = SQL & " CASE (convert(decimal(20,8),ISNULL(( "
SQL = SQL & " SUM(CASE WHEN Resultado = 'Chamadas' THEN Qtde ELSE 0 END) "
SQL = SQL & "  ),0))) WHEN 0 THEN 0 ELSE ISNULL(round((convert(decimal(20,8),ISNULL( "
SQL = SQL & " SUM(CASE WHEN Resultado = 'Marcacoes' THEN Qtde ELSE 0 END) "
SQL = SQL & " ,0))) / (convert(decimal(20,8),ISNULL( "
SQL = SQL & " SUM(CASE WHEN Resultado = 'Chamadas' THEN Qtde ELSE 0 END) "
SQL = SQL & " ,0))) * 100,2),0) END AS Tx_Marcacao, "

SQL = SQL & " CASE (convert(decimal(20,8),ISNULL(( "
SQL = SQL & " (CASE SUM(CASE WHEN Resultado = 'Chamadas' THEN Qtde ELSE 0 END) "
SQL = SQL & " WHEN 0 THEN 0 ELSE "
SQL = SQL & " SUM(CASE WHEN Resultado = 'TEMPO' THEN Qtde ELSE 0 END) / "
SQL = SQL & " SUM(CASE WHEN Resultado = 'Chamadas' THEN Qtde ELSE 0 END) "
SQL = SQL & " END) "
SQL = SQL & " ),0))) WHEN 0 THEN 0 ELSE ISNULL(round((convert(decimal(20,8),ISNULL( "
SQL = SQL & " (CASE (convert(decimal(20,8),ISNULL(( "
SQL = SQL & " SUM(CASE WHEN Resultado = 'Chamadas' THEN Qtde ELSE 0 END) "
SQL = SQL & "  ),0))) WHEN 0 THEN 0 ELSE ISNULL(round((convert(decimal(20,8),ISNULL( "
SQL = SQL & " SUM(CASE WHEN Resultado = 'Pedidos' THEN Qtde ELSE 0 END) "
SQL = SQL & " ,0))) / (convert(decimal(20,8),ISNULL( "
SQL = SQL & " SUM(CASE WHEN Resultado = 'Chamadas' THEN Qtde ELSE 0 END) "
SQL = SQL & " ,0))) * 100,2),0) END) "
SQL = SQL & " ,0))) /  "
SQL = SQL & " (convert(decimal(20,8),ISNULL( "
SQL = SQL & " (CASE SUM(CASE WHEN Resultado = 'Chamadas' THEN Qtde ELSE 0 END) "
SQL = SQL & " WHEN 0 THEN 0 ELSE "
SQL = SQL & " SUM(CASE WHEN Resultado = 'TEMPO' THEN Qtde ELSE 0 END) / "
SQL = SQL & " SUM(CASE WHEN Resultado = 'Chamadas' THEN Qtde ELSE 0 END) "
SQL = SQL & " END) "
SQL = SQL & " ,0))) ,3),0) END AS Fator "

SQL = SQL & " from (  "
SQL = SQL & " SELECT     U.usuario, CE.datetime_entry_queue, COUNT(CE.status) AS Qtde, 'Chamadas' as Resultado "
SQL = SQL & " FROM PABX...agent AS A  "
SQL = SQL & " LEFT OUTER JOIN PABX...call_entry AS CE ON A.id = CE.id_agent  "
SQL = SQL & " RIGHT OUTER JOIN SIS_usuarios AS U ON A.number = U.usuario_telefonia "
SQL = SQL & " WHERE id_queue_call_entry in ('1', '2') "
SQL = SQL & " GROUP BY U.usuario, CE.datetime_entry_queue "
SQL = SQL & " UNION ALL "
SQL = SQL & " SELECT     U.usuario, CE.datetime_entry_queue, SUM(CE.Duration) AS Qtde, 'TEMPO' as Resultado "
SQL = SQL & " FROM PABX...agent AS A  "
SQL = SQL & " LEFT OUTER JOIN PABX...call_entry AS CE ON A.id = CE.id_agent  "
SQL = SQL & " RIGHT OUTER JOIN SIS_usuarios AS U ON A.number = U.usuario_telefonia "
SQL = SQL & " WHERE id_queue_call_entry in ('1', '2') "
SQL = SQL & " GROUP BY U.usuario, CE.datetime_entry_queue "
SQL = SQL & " UNION ALL "
SQL = SQL & " SELECT usuario, data_producao, COUNT(fa.ID) AS Qtde, 'Pedidos' as Resultado "
SQL = SQL & " FROM tbPedidosListo AS FA INNER JOIN SIS_usuarios AS U ON FA.id_usuario_vendedor = U.usuario_listo "
SQL = SQL & " WHERE (FA.status = N'7') "
SQL = SQL & " GROUP BY usuario, data_producao "
SQL = SQL & " UNION ALL "
SQL = SQL & " SELECT     ResponsavelCriacao, DataCriacao, COUNT(ID) AS Qtde, 'Marcacoes' as Resultado "
SQL = SQL & " FROM         tbFichasAtendimento AS FA "
SQL = SQL & " GROUP BY ResponsavelCriacao, DataCriacao "

SQL = SQL & " ) CONSOLIDADO "
If Len(Request.QueryString("Data")) < 8 Then
SQL = SQL & " WHERE MONTH(datetime_entry_queue) = " & Mes & " AND YEAR(datetime_entry_queue) = " & Ano & " "
Else
	SQL = SQL & " Where datetime_entry_queue between CONVERT(DATETIME,'"&Ano&"-"&Mes&"-"&Dia&" 00:00:00',102) and CONVERT(DATETIME,'"&Ano&"-"&Mes&"-"&Dia&" 23:59:00',102)"
End If
If Vendedor <> "0" Then
	SQL = SQL & " AND (usuario = N'"&Vendedor&"') "
End If
	SQL = SQL & " AND (usuario not in ('admin', 'myrlla', 'marina.silva', 'ruslana.pires')) "
SQL = SQL & " GROUP BY Usuario Order by Fator DESC "

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

	Marcacoes = RSBUSCAS("Marcacoes")
	Marcacoes = FormatNumber(Marcacoes,0,-1,0,-2)
	Pedido = RSBUSCAS("Pedido")
	Pedido = FormatNumber(Pedido,0,-1,0,-2)
	Chamadas = RSBUSCAS("Chamadas")
	Chamadas = FormatNumber(Chamadas,0,-1,0,-2)
	TMA = RSBUSCAS("TMA")
	TMA = FormatNumber(TMA,0,-1,0,-2)
	TEMPO = RSBUSCAS("TEMPO")
	TEMPO = FormatNumber(TEMPO,0,-1,0,-2)
	Fator = RSBUSCAS("Fator")
	Fator = FormatNumber(Fator,3,-1,0,-2)

	If Chamadas > 0 Then
				Tx_Conversao = (CDbl(Pedido))/ (CDbl(Chamadas) )*100
	Else
	Tx_Conversao = 0
	End If
	Tx_Conversao = FormatNumber(Tx_Conversao,2,-1,0,-2)

	If Chamadas > 0 Then
				Tx_Marcacao = (CDbl(Marcacoes))/ (CDbl(Chamadas) )*100
	Else
	Tx_Marcacao = 0
	End If
	Tx_Marcacao = FormatNumber(Tx_Marcacao,2,-1,0,-2)

	Marcacoes_TOTAL = Marcacoes_TOTAL + CDbl(Marcacoes)
	Pedido_TOTAL = Pedido_TOTAL + CDbl(Pedido)
	Chamadas_TOTAL = Chamadas_TOTAL + CDbl(Chamadas)
	Tempo_TOTAL = Tempo_TOTAL + CDbl(Tempo)

	If Chamadas_TOTAL > 0 Then
				Tx_Conversao_Total = (CDbl(Pedido_TOTAL))/ (CDbl(Chamadas_TOTAL) )*100
	Else
	Tx_Conversao_Total = 0
	End If
	Tx_Conversao_Total = FormatNumber(Tx_Conversao_Total,2,-1,0,-2)

	If Chamadas_TOTAL > 0 Then
				Tx_Marcacao_Total = (CDbl(Marcacoes_TOTAL))/ (CDbl(Chamadas_TOTAL) )*100
	Else
	Tx_Marcacao_Total = 0
	End If
	Tx_Marcacao_Total = FormatNumber(Tx_Marcacao_Total,2,-1,0,-2)	

	If Chamadas_TOTAL > 0 Then
				TMA_Total = (CDbl(TEMPO_TOTAL))/ (CDbl(Chamadas_TOTAL) )
	Else
	TMA_Total = 0
	End If
	TMA_Total = FormatNumber(TMA_Total,0,-1,0,-2)

If i = "1" Then Bonus = "" End If
If i = "2" Then Bonus = "" End If
If i = "3" Then Bonus = "" End If
If i = "4" Then Bonus = "" End If

If Tx_Conversao > 60 Then
IndConversao = "BallGreen"
End If
If Tx_Conversao = 60 Then
IndConversao = "BallYellow"
End If
If Tx_Conversao < 60 Then
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

If Pedido_TOTAL >= 8000 Then
Vlr_1Quartil = 0.50
Vlr_2Quartil = 0.30
Vlr_3Quartil = 0.20
Vlr_4Quartil = 0.05
Else
Vlr_1Quartil = 0.3
Vlr_2Quartil = 0.25
Vlr_3Quartil = 0.15
Vlr_4Quartil = 0.05
End If

If Quartil = "1" Then
Comissao = (Pedido*Vlr_1Quartil)
End If
If Quartil = "2" Then
Comissao = (Pedido*Vlr_2Quartil)
End If
If Quartil = "3" Then
Comissao = (Pedido*Vlr_3Quartil)
End If
If Quartil = "4" Then
Comissao = (Pedido*Vlr_4Quartil)
End If
	Comissao = FormatNumber(Comissao,2,-1,0,-2)
	Comissao_TOTAL = Comissao_TOTAL + CDbl(Comissao)
	Comissao_TOTAL = FormatNumber(Comissao_TOTAL,2,-1,0,-2)

min = TMA \ 60
segundos = TMA MOD 60

min_Total = TMA_Total \ 60
segundos_Total = TMA_Total MOD 60


%>
 <tr>
    <td width="5%" align="left"><div align="center"><%=i%></div></td>
    <td width="5%" align="left"><div align="center"><%=Quartil%></div></td>    
    <td width="25%" align="left"><%=RSBuscas("usuario")%></td>
    <td width="10%" align="left"><div align="center"><%=Tx_Marcacao%>%</td>    
    <td width="10%" align="center"><div align="center"><%=RSBuscas("Pedido")%> / <%=RSBuscas("Chamadas")%></div></td>
    <td width="10%" align="center"><div align="center"><%=Tx_Conversao%>% <img src="../imagens/<%=IndConversao%>.gif" width="15" height="18" border="0"></div></td>
    <td width="10%" align="center"><div align="center"><%=TMA%> (<%=min%>:<%=segundos%>)</div></td>
    <td width="10%" align="center"><div align="center"><%=Fator%></div></td>
<%If Acesso <> "Vendas" Then %>
    <td width="10%" align="center"><div align="center">R$ <%=Comissao%>  </div></td>       
<%End If%>
  </tr>
  <%
	RSBUSCAS.Movenext
Loop
%>
<tr bgcolor="#cdd5da">
    <td colspan="3" align="left"><div align="left"><b>Total</b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=Tx_Marcacao_Total%>%</b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=Pedido_Total%> / <%=Chamadas_Total%></b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=Tx_Conversao_Total%>%</b></div></td>
	<td width="10%" align="center"><div align="center"><b><%=TMA_Total%> (<%=min_Total%>:<%=segundos_Total%>)</b></div></td>
	<td width="10%" align="center"><div align="center"><b>-</b></div></td>
<%If Acesso <> "Vendas" Then %>
	<td width="10%" align="center"><div align="center"><b>R$ <%=Comissao_Total%></b></div></td>         
<% End If %>    
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

