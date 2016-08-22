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
<title>DNA</title>
<META HTTP-EQUIV="Refresh" CONTENT="300;URL="tempoRealAgendamento_Painel.asp?Acumulado=<%=Acumulado%>&Diario=<%=Diario%>&Enviar=S">
<link rel="stylesheet" href="../include/pgo.css" type="text/css">
<script language="javascript" src="../Include/MostraCalendario.js"></script>
<script language="javascript">
	function ValidaDados() {
		if (confirm("Confirma a busca?")) {
				frmDados.action = "tempoRealAgendamento_Painel.asp?Acumulado="+frmDados.txtAcumulado.value+"&Diario="+frmDados.txtDiario.value+"&Enviar=S";
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
      <td bgcolor="#EBF3F1"> <p align="center"><b>Acompanhamento Tempo Real - NET Agendamento</b></td>
    </tr>
    <tr>
      <td><p align="center"><i>Quantidade de Agendamento Acumuladas </i>: <input name="txtAcumulado" type="text" class="formfield" size="15" maxlength="10" readonly>
      </td>
    </tr>
    <tr>
      <td><p align="center"><i>Meta HOJE:</i> <input name="txtDiario" type="text" class="formfield" size="15" maxlength="10" readonly>
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

	SQL2 = " SELECT MAX(convert(datetime, data, 103)) as Data FROM NET_ATOM_Protocolo_Tabulacao_Agendamento"
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
    <td colspan="9"> <p align="center"><b>Acompanhamento ATOM - NET Agendamento (Atualização ATOM até: <%=Maximo%>)</b></td>
  </tr>
  <!--#include file="AbreConexao.asp"-->
<%
Data = Year(Data)&"-"&Month(Data)&"-"&Day(Data)
If Data = "" Then
	Data = Year(now)&"-"&Month(now)&"-"&Day(now)
End If

If weekday(now()) = 1 Then Diario = 700 End If
If weekday(now()) = 2 Then Diario = 700 End If
If weekday(now()) = 3 Then Diario = 700 End If
If weekday(now()) = 4 Then Diario = 700 End If
If weekday(now()) = 5 Then Diario = 700 End If
If weekday(now()) = 6 Then Diario = 700 End If
If weekday(now()) = 7 Then Diario = 700 End If

SQL = " SELECT "
SQL = SQL & " SUM(Agendamento) as Agendamento, "
SQL = SQL & " SUM(Atendida) as Atendidas, "
SQL = SQL & " SUM(Telefonia) as Telefonia, "
SQL = SQL & " SUM(Backlog) as Backlog, "
SQL = SQL & " SUM(Agendamento)+ "
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
SQL = SQL & " WHERE "
SQL = SQL & " DAY(convert(datetime, data, 103)) = DAY(GETDATE()) AND "
SQL = SQL & " MONTH(convert(datetime, data, 103)) = MONTH(GETDATE()) AND  "
SQL = SQL & " YEAR(convert(datetime, data, 103)) = YEAR(GETDATE()) "
SQL = SQL & " GROUP BY [Nome Operador], SRAG.Agendamento, SRAG.Atendida, SRAG.Telefonia ) A "


Set RSBUSCAS = server.createobject("ADODB.Recordset")
'RESPONSE.WRITE SQL
RSBUSCAS.Open SQL, Conexao
i = 0
Do While Not RSBUSCAS.EOF

Agendamento = RSBUSCAS("Agendamento")
Agendamento = FormatNumber(Agendamento,0,-1,0,-2)

Atendidas = RSBUSCAS("Atendidas")
Atendidas = FormatNumber(Atendidas,0,-1,0,-2)

Telefonia = RSBUSCAS("Telefonia")
Telefonia = FormatNumber(Telefonia,0,-1,0,-2)

Total = RSBUSCAS("Total")
Total = FormatNumber(Total,0,-1,0,-2)

TxAgendamento = RSBUSCAS("TxAgendamento")
TxAgendamento = FormatNumber(TxAgendamento,2,-1,0,-2)

TxImprodImpossivel = RSBUSCAS("TxImprodImpossivel")
TxImprodImpossivel = FormatNumber(TxImprodImpossivel,2,-1,0,-2)

TxImprodPossivel = RSBUSCAS("TxImprodPossivel")
TxImprodPossivel = FormatNumber(TxImprodPossivel,2,-1,0,-2)

TxAtend = RSBUSCAS("TxAtend")
TxAtend = FormatNumber(TxAtend,2,-1,0,-2)

%>
</table>

<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr>
    <td width="25%" align="center"><p align="center" class="LetrasGrandes" >%Agend: <br><%=TxAgendamento%>%<br>AG.: <%=Agendamento%></td>
    <td width="25%" align="center"><p align="center" class="LetrasGrandes" >Sem Possib.: <br><%=TxImprodImpossivel%>%</td>
	<td width="25%" align="center"><p align="center" class="LetrasGrandes" >Sem Sucesso: <br><%=TxImprodPossivel%>%</td>
	<td width="25%" align="center"><p align="center" class="LetrasGrandes" >Atend: <br><%=TxAtend%>%</td>
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
    <td colspan="9"> <p align="center"><b>Ranking Agendamento do dia <%=now%></b></td>
  </tr>
 
  <tr bgcolor="#EBF3F1">
    <td width="20%" align="center"><p align="center"><i>Intervalo</i></td>
    <td width="10%" align="center"><div align="center">Agendamento</div></td>
    <td width="10%" align="center"><div align="center">Atendidas</div></td>
    <td width="10%" align="center"><div align="center">Telefonia</div></td>
    <td width="10%" align="center"><div align="center">Tx. Agend.</div></td>
	<td width="10%" align="center"><div align="center">Sem Sucesso</div></td>
	<td width="10%" align="center"><div align="center">Sem Possibilidade</div></td>
    <td width="10%" align="center"><div align="center">Tx. Atend.</div></td>
      </tr>
<!--#include file="AbreConexao.asp",,,,,,,-->
<%

SQL = " SELECT "

SQL = SQL & " case "
SQL = SQL & " when cast(right(left((Convert(VarChar(8), convert(datetime, data , 103), 108)),5),2) as Int) < 30 then left((Convert(VarChar(8), convert(datetime, data , 103), 108)),2) "
SQL = SQL & " + ':00' "
SQL = SQL & " + ' Até ' "
SQL = SQL & " + left((Convert(VarChar(8), convert(datetime, data , 103), 108)),2) + ':30' "
SQL = SQL & " when cast(right(left((Convert(VarChar(8), convert(datetime, data , 103), 108)),5),2) as Int) >= 30 then left((Convert(VarChar(8), convert(datetime, data , 103), 108)),2) + ':30' + ' Até ' "
SQL = SQL & " + right('0' + cast(cast(left((Convert(VarChar(8), convert(datetime, data , 103), 108)),2) as int) + 1 as varchar),2) +':00' end as Intraday,  "
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
SQL = SQL & " SELECT convert(datetime, TAG.DATA, 103) as Data,  "
SQL = SQL & " SUM(cast(isnull(SRAG.Agendamento,0) as int)) as Agendamento,  "
SQL = SQL & " SUM(cast(isnull(SRAG.Atendida,0) as int)) as Atendida,  "
SQL = SQL & " SUM(cast(isnull(SRAG.Telefonia,0) as int)) as Telefonia,  "
SQL = SQL & " SUM(cast(isnull(SRAG.ImprodPossivel,0) as int)) as ImprodPossivel,  "
SQL = SQL & " SUM(cast(isnull(SRAG.ImprodImpossivel,0) as int)) as ImprodImpossivel,  "
SQL = SQL & " SUM(cast(isnull(SRAG.Backlog,0) as int)) as Backlog,  "
SQL = SQL & " COUNT(*) as Qtde  "
SQL = SQL & " FROM NET_ATOM_Protocolo_Tabulacao_Agendamento TAG "
SQL = SQL & " LEFT OUTER JOIN NET_ATOMAgendamento_SISResultado SRAG ON TAG.Resultado = SRAG.Resultado "
SQL = SQL & " WHERE "
SQL = SQL & " DAY(convert(datetime, data, 103)) = DAY(GETDATE()) AND "
SQL = SQL & " MONTH(convert(datetime, data, 103)) = MONTH(GETDATE()) AND  "
SQL = SQL & " YEAR(convert(datetime, data, 103)) = YEAR(GETDATE()) "
SQL = SQL & " AND [Nome Operador] NOT LIKE '%OPERADOR DISCADOR%' "
SQL = SQL & " GROUP BY Data ) A "
SQL = SQL & " GROUP BY "
SQL = SQL & " case "
SQL = SQL & " when cast(right(left((Convert(VarChar(8), convert(datetime, data , 103), 108)),5),2) as Int) < 30 then left((Convert(VarChar(8), convert(datetime, data , 103), 108)),2) "
SQL = SQL & " + ':00' "
SQL = SQL & " + ' Até ' "
SQL = SQL & " + left((Convert(VarChar(8), convert(datetime, data , 103), 108)),2) + ':30' "
SQL = SQL & " when cast(right(left((Convert(VarChar(8), convert(datetime, data , 103), 108)),5),2) as Int) >= 30 then left((Convert(VarChar(8), convert(datetime, data , 103), 108)),2) + ':30' + ' Até ' "
SQL = SQL & " + right('0' + cast(cast(left((Convert(VarChar(8), convert(datetime, data , 103), 108)),2) as int) + 1 as varchar),2) +':00' end  "
SQL = SQL & " ORDER BY  "
SQL = SQL & " case "
SQL = SQL & " when cast(right(left((Convert(VarChar(8), convert(datetime, data , 103), 108)),5),2) as Int) < 30 then left((Convert(VarChar(8), convert(datetime, data , 103), 108)),2) "
SQL = SQL & " + ':00' "
SQL = SQL & " + ' Até ' "
SQL = SQL & " + left((Convert(VarChar(8), convert(datetime, data , 103), 108)),2) + ':30' "
SQL = SQL & " when cast(right(left((Convert(VarChar(8), convert(datetime, data , 103), 108)),5),2) as Int) >= 30 then left((Convert(VarChar(8), convert(datetime, data , 103), 108)),2) + ':30' + ' Até ' "
SQL = SQL & " + right('0' + cast(cast(left((Convert(VarChar(8), convert(datetime, data , 103), 108)),2) as int) + 1 as varchar),2) +':00' end ASC  "

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
  
    <td width="20%" align="left"><%=RSBuscas("Intraday")%></td>
    <td width="10%" align="left"><div align="center"><%= 	%></td>    
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
    <td colspan="1" align="left"><div align="left"><b>Total</b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=Agendamento_Total%></b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=Atendidas_Total%></b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=Telefonia_Total%></b></div></td>
	<td width="10%" align="center"><div align="center"><b><%=TxAgendamento_Total%>%</b></div></td>
	<td width="10%" align="center"><div align="center"><b><%=TxImprodPossivel_Total%>%</b></div></td>
	<td width="10%" align="center"><div align="center"><b><%=TxImprodImpossivel_Total%>%</b></div></td>
	<td width="10%" align="center"><div align="center"><b><%=TxAtend_Total%>%</b></div></td>
 
  </tr> 
 

</table>
    <input type="hidden" name="txtAcumulado" value="<%=Acumulado%>">
    <input type="hidden" name="txtDiario" value="<Diario%>">

</body>
</html>