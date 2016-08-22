<%
Dim i, Data, Enviar, ID, Acumulado, Diario, Vlr_1Quartil, Vlr_2Quartil, Vlr_3Quartil, Vlr_4Quartil
Enviar = Request.QueryString("Enviar")
ID = Request.QueryString("ID")
Acumulado = Request.QueryString("Acumulado")
Diario = Request.QueryString("Diario")
User = Trim(Session("usuario"))
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>WeDo Servi�os</title>
<META HTTP-EQUIV="Refresh" CONTENT="30;URL=tempoRealPedidos_comissao.asp?Acumulado=<%=Acumulado%>&Diario=<%=Diario%>&Enviar=S">
<link rel="stylesheet" href="../include/pgo.css" type="text/css">
<script language="javascript" src="../Include/MostraCalendario.js"></script>
<script language="javascript">
	function ValidaDados() {
		if (confirm("Confirma a busca?")) {
				frmDados.action = "tempoRealPedidos_comissao.asp?Acumulado="+frmDados.txtAcumulado.value+"&Diario="+frmDados.txtDiario.value+"&Enviar=S";
				frmDados.submit();
		}
		return false;
	}
</script>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="javascript:document.frmDados.txtDDDTelefone.focus();">
<form method="POST" name="frmDados">
    <input type="hidden" name="txtUsuario" value="<%=User%>">
<% If Enviar <> "S" Then %>
  <table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center" >
    <tr bgcolor="#EBF3F1">
      <td bgcolor="#EBF3F1"> <p align="center"><b>Acompanhamento Tempo Real - Pedidos</b></td>
    </tr>
    <tr>
      <td><p align="center"><i>Quantidade de Pedidos Acumulados (Listo)</i>: <input name="txtAcumulado" type="text" class="formfield" size="15" maxlength="10">
      </td>
    </tr>
    <tr>
      <td><p align="center"><i>Meta HOJE:</i> <input name="txtDiario" type="text" class="formfield" size="15" maxlength="10">
      </td>
    </tr>
    <tr>
      <td> <p align="center">
          <input name="Buscar" type="submit" onClick="return ValidaDados();" id="Buscar" value="Buscar" class="formbutton">
      </td>
    </tr>
  </table>
<% End if %>  
</form>
  <!--#include file="AbreConexao.asp"-->
<%
Sub Consultar

	SQL2 = " SELECT MAX(DataCriacao) AS DATA FROM tbFichasAtendimento "
	Set RSMAXIMO = server.createobject("ADODB.Recordset")
	RSMAXIMO.Open SQL2, Conexao
	If Not RSMAXIMO.EOF Then
		Maximo = Trim(RSMAXIMO("DATA"))
	End If

Set RSMAXIMO = Nothing
Set Conexao = Nothing


%>

<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td colspan="9"> <p align="center"><b>Falta Para a Meta (�ltimo Pedido: <%=Maximo%>)</b></td>
  </tr>
  <!--#include file="AbreConexao.asp"-->
<%
Data = Year(Data)&"-"&Month(Data)&"-"&Day(Data)
If Data = "" Then
	Data = Year(now)&"-"&Month(now)&"-"&Day(now)
End If

SQL = " SELECT   10000 - "& Acumulado &" - COUNT(*) AS Qtde, "& Diario &" - COUNT(*) as QtdeDiario  "
SQL = SQL & " FROM         tbFichasAtendimento "
SQL = SQL & " WHERE     (DetResultado = N'Pedido Realizado') AND"
SQL = SQL & " Day(DataCriacao) = Day(getdate()) and MONTH(DataCriacao) = month(getdate()) AND YEAR(DataCriacao) = year(getdate()) "

Set RSBUSCAS = server.createobject("ADODB.Recordset")
'RESPONSE.WRITE SQL
RSBUSCAS.Open SQL, Conexao
i = 0
Do While Not RSBUSCAS.EOF

PedidosMes = RSBUSCAS("Qtde")
PedidosMes = FormatNumber(PedidosMes,0,-1,0,-2)

PedidosDiario = RSBUSCAS("QtdeDiario")
PedidosDiario = FormatNumber(PedidosDiario,0,-1,0,-2)

If Acumulado >= "10000" Then
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

%>
</table>

<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr>
    <td width="8%" align="center"><p align="center" class="LetrasGrandes" > M�s: <%=PedidosMes%><br>Dia: <%=PedidosDiario%></td>
   </tr>
  <%
  Response.Flush
  i=i+1
	RSBUSCAS.Movenext
Loop
%>
  <%
End Sub
If Enviar = "S"  Then
	Consultar
End If
%>
</table>

    <td width="33%" rowspan="2" align="center" valign="top"><table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td colspan="10"> <p align="center"><b>Ranking Vendas do dia <%=now%></b></td>
  </tr>
  <tr bgcolor="#EBF3F1">
    <td width="5%" align="center"><i>Rank</i></td>
    <td width="5%" align="center"><i>Quartil</i></td>
    <td width="25%" align="center"><p align="center"><i>Nome</i></td>
    <td width="10%" align="center"><div align="center">% Marca��o DNA</div></td>
    <td width="10%" align="center"><div align="center">Pedidos / Liga��es</div></td>
    <td width="10%" align="center"><div align="center">Tx. Convers�o</div></td>
    <td width="10%" align="center"><div align="center">TMA(s)</div></td>
    <td width="10%" align="center"><div align="center">Nota</div></td>
    <td width="10%" align="center"><div align="center">Comiss�o</div></td>    
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
SQL = SQL & " WHERE Day(datetime_entry_queue) = day(getdate()) and MONTH(datetime_entry_queue) = month(getdate()) AND YEAR(datetime_entry_queue) = year(getdate()) "
SQL = SQL & " GROUP BY U.usuario, CE.datetime_entry_queue "
SQL = SQL & " UNION ALL "
SQL = SQL & " SELECT     U.usuario, CE.datetime_entry_queue, SUM(CE.Duration) AS Qtde, 'TEMPO' as Resultado "
SQL = SQL & " FROM PABX...agent AS A  "
SQL = SQL & " LEFT OUTER JOIN PABX...call_entry AS CE ON A.id = CE.id_agent  "
SQL = SQL & " RIGHT OUTER JOIN SIS_usuarios AS U ON A.number = U.usuario_telefonia "
SQL = SQL & " WHERE Day(datetime_entry_queue) = day(getdate()) and MONTH(datetime_entry_queue) = month(getdate()) AND YEAR(datetime_entry_queue) = year(getdate()) "
SQL = SQL & " GROUP BY U.usuario, CE.datetime_entry_queue "
SQL = SQL & " UNION ALL "
SQL = SQL & " SELECT     ResponsavelCriacao, DataCriacao, COUNT(ID) AS Qtde, 'Pedidos' as Resultado "
SQL = SQL & " FROM         tbFichasAtendimento AS FA "
SQL = SQL & " WHERE     (ResultadoChamada = N'Pedido') and"
SQL = SQL & " Day(DataCriacao) = day(getdate()) and MONTH(DataCriacao) = month(getdate()) AND YEAR(DataCriacao) = year(getdate()) "
SQL = SQL & " GROUP BY ResponsavelCriacao, DataCriacao  "
SQL = SQL & " UNION ALL "
SQL = SQL & " SELECT     ResponsavelCriacao, DataCriacao, COUNT(ID) AS Qtde, 'Marcacoes' as Resultado "
SQL = SQL & " FROM         tbFichasAtendimento AS FA "
SQL = SQL & " WHERE Day(DataCriacao) = day(getdate()) and MONTH(DataCriacao) = month(getdate()) AND YEAR(DataCriacao) = year(getdate()) "
SQL = SQL & " GROUP BY ResponsavelCriacao, DataCriacao "
SQL = SQL & " ) As Consolidado "
SQL = SQL & " GROUP BY Usuario Order by Fator DESC "
'response.write SQL
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
    <td width="10%" align="center"><div align="center">R$ <%=Comissao%>  </div></td>       
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
	<td width="10%" align="center"><div align="center"><b>-</b></div></td>         
  </tr>
</table>
    <input type="hidden" name="txtAcumulado" value="<%=Acumulado%>">
    <input type="hidden" name="txtDiario" value="<Diario%>">
</body>
</html>