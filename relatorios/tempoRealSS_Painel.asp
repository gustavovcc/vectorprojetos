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
<META HTTP-EQUIV="Refresh" CONTENT="30;URL="tempoRealSS_Painel.asp?Acumulado=<%=Acumulado%>&Diario=<%=Diario%>&Enviar=S">
<link rel="stylesheet" href="../include/pgo.css" type="text/css">
<script language="javascript" src="../Include/MostraCalendario.js"></script>
<script language="javascript">
	function ValidaDados() {
		if (confirm("Confirma a busca?")) {
				frmDados.action = "tempoRealSS_Painel.asp?Acumulado="+frmDados.txtAcumulado.value+"&Diario="+frmDados.txtDiario.value+"&Enviar=S";
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
      <td bgcolor="#EBF3F1"> <p align="center"><b>Acompanhamento Tempo Real - Serviços NET</b></td>
    </tr>
    <tr>
      <td><p align="center"><i>Quantidade de OS Acumuladas (Imperium)</i>: <input name="txtAcumulado" type="text" class="formfield" size="15" maxlength="10" readonly>
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

	SQL2 = " SELECT MAX(DATABaixa) as Data FROM NET_IMPERIUM"
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
    <td colspan="9"> <p align="center"><b>Acompanhamento SS - NET (Atualizado até: <%=Maximo%>)</b></td>
  </tr>
  <!--#include file="AbreConexao.asp"-->
<%
Data = Year(Data)&"-"&Month(Data)&"-"&Day(Data)
If Data = "" Then
	Data = Year(now)&"-"&Month(now)&"-"&Day(now)
End If

If weekday(now()) = 1 Then Diario = 639 End If
If weekday(now()) = 2 Then Diario = 154 End If
If weekday(now()) = 3 Then Diario = 164 End If
If weekday(now()) = 4 Then Diario = 231 End If
If weekday(now()) = 5 Then Diario = 273 End If
If weekday(now()) = 6 Then Diario = 321 End If
If weekday(now()) = 7 Then Diario = 455 End If


SQL = SQL & " SELECT  "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVSucesso' THEN Qtde ELSE 0 END) as SrvSucesso, "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVInsucesso' THEN Qtde ELSE 0 END) as SrvInSucesso, "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVPendente' THEN Qtde ELSE 0 END) as SrvPendente, "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTSucesso' THEN Qtde ELSE 0 END) as VTSucesso, "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTInsucesso' THEN Qtde ELSE 0 END) as VTInSucesso, "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTPendente' THEN Qtde ELSE 0 END) as VTPendente, "

SQL = SQL & " CASE (convert(decimal(20,8),ISNULL((  "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVSucesso' THEN Qtde ELSE 0 END)+ "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVInsucesso' THEN Qtde ELSE 0 END)+ "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVPendente' THEN Qtde ELSE 0 END) "
SQL = SQL & " ),0))) WHEN 0 THEN 0 ELSE ISNULL(round((convert(decimal(20,8),ISNULL(  "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVSucesso' THEN Qtde ELSE 0 END) "
SQL = SQL & " ,0))) / (convert(decimal(20,8),ISNULL(  "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVSucesso' THEN Qtde ELSE 0 END)+ "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVInsucesso' THEN Qtde ELSE 0 END)+ "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVPendente' THEN Qtde ELSE 0 END) "
SQL = SQL & " ,0))) * 100,2),0) END AS Tx_SucessoSRV, "

SQL = SQL & " CASE (convert(decimal(20,8),ISNULL((  "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTSucesso' THEN Qtde ELSE 0 END)+ "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTInsucesso' THEN Qtde ELSE 0 END)+ "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTPendente' THEN Qtde ELSE 0 END) "
SQL = SQL & " ),0))) WHEN 0 THEN 0 ELSE ISNULL(round((convert(decimal(20,8),ISNULL(  "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTSucesso' THEN Qtde ELSE 0 END) "
SQL = SQL & " ,0))) / (convert(decimal(20,8),ISNULL(  "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTSucesso' THEN Qtde ELSE 0 END)+ "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTInsucesso' THEN Qtde ELSE 0 END)+ "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTPendente' THEN Qtde ELSE 0 END) "
SQL = SQL & " ,0))) * 100,2),0) END AS Tx_SucessoVT "

SQL = SQL & " FROM ( "
SQL = SQL & " SELECT     'SRV'+isnull(IDP.Status,'Pendente') AS Resultado, COUNT(*) AS Qtde "
SQL = SQL & " FROM          "
SQL = SQL & " NET_IMPERIUM AS I LEFT OUTER JOIN "
SQL = SQL & " NET_IMPERIUMDeParaBaixa AS IDP ON I.CodigoBaixa = IDP.Codigo "
SQL = SQL & " WHERE      "
SQL = SQL & " DAY(I.Data) = DAY(GETDATE()) AND "
SQL = SQL & " MONTH(I.Data) = MONTH(GETDATE()) AND  "
SQL = SQL & " YEAR(I.Data) = YEAR(GETDATE()) "
SQL = SQL & " AND Turno NOT LIKE '%VT%' "
SQL = SQL & " GROUP BY IDP.Status "
SQL = SQL & " UNION ALL "
SQL = SQL & " SELECT     'VT'+isnull(IDP.StatusVT,'Pendente') AS Resultado, COUNT(*) AS Qtde "
SQL = SQL & " FROM          "
SQL = SQL & " NET_IMPERIUM AS I LEFT OUTER JOIN "
SQL = SQL & " NET_IMPERIUMDeParaBaixa AS IDP ON I.CodigoBaixa = IDP.Codigo "
SQL = SQL & " WHERE "     
SQL = SQL & " DAY(I.Data) = DAY(GETDATE()) AND "
SQL = SQL & " MONTH(I.Data) = MONTH(GETDATE()) AND "
SQL = SQL & " YEAR(I.Data) = YEAR(GETDATE()) AND "
SQL = SQL & " Turno LIKE '%VT%' "
SQL = SQL & " GROUP BY IDP.StatusVT ) AS CONSOLIDADO "

Set RSBUSCAS = server.createobject("ADODB.Recordset")
'RESPONSE.WRITE SQL
RSBUSCAS.Open SQL, Conexao
i = 0
Do While Not RSBUSCAS.EOF

SrvSucesso = RSBUSCAS("SrvSucesso")
SrvSucesso = FormatNumber(SrvSucesso,0,-1,0,-2)

SrvInsucesso = RSBUSCAS("SrvInsucesso")
SrvInsucesso = FormatNumber(SrvInsucesso,0,-1,0,-2)

SrvPendente = RSBUSCAS("SrvPendente")
SrvPendente = FormatNumber(SrvPendente,0,-1,0,-2)

Tx_SucessoSRV = RSBUSCAS("Tx_SucessoSRV")
Tx_SucessoSRV = FormatNumber(Tx_SucessoSRV,0,-1,0,-2)

VTSucesso = RSBUSCAS("VTSucesso")
VTSucesso = FormatNumber(VTSucesso,0,-1,0,-2)

VTInsucesso = RSBUSCAS("VTInsucesso")
VTInsucesso = FormatNumber(VTInsucesso,0,-1,0,-2)

VTPendente = RSBUSCAS("VTPendente")
VTPendente = FormatNumber(VTPendente,0,-1,0,-2)

Tx_SucessoVT = RSBUSCAS("Tx_SucessoVT")
Tx_SucessoVT = FormatNumber(Tx_SucessoVT,0,-1,0,-2)

%>
</table>

<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr>
    <td width="8%" align="center"><p align="right" class="LetrasGrandes" > %Suces. Serv.: <%=Tx_SucessoSRV%>%<br>Sucesso Serviço: <%=SrvSucesso%><br>Insucesso Serviço: <%=SrvInsucesso%><br>Serv. Pendentes: <%=SrvPendente%></td>
    <td width="8%" align="center"><p align="right" class="LetrasGrandes" > %Suces. VT.: <%=Tx_SucessoVT%>%<br>Sucesso VT: <%=VTSucesso%><br>Insucesso VT: <%=VTInsucesso%><br>VT Pendentes: <%=VTPendente%></td>
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
    <td colspan="11"> <p align="center"><b>Ranking Execução do dia <%=now%></b></td>
  </tr>
  <tr bgcolor="#EBF3F1">
    <td width="5%" align="center"><i>Rank</i></td>
    <td width="5%" align="center"><i>Quartil</i></td>
    <td width="10%" align="center"><p align="center"><i>Nome</i></td>
    <td width="10%" align="center"><div align="center">Serviço<br>Sucesso</div></td>
    <td width="10%" align="center"><div align="center">Serviço<br>Insucesso</div></td>
    <td width="10%" align="center"><div align="center">Serviço<br>Pendente</div></td>
    <td width="10%" align="center"><div align="center">Tx. Sucesso SRV<br>Pendente</div></td>
    <td width="10%" align="center"><div align="center">VT<br>Sucesso</div></td>
    <td width="10%" align="center"><div align="center">VT<br>Insucesso</div></td>
    <td width="10%" align="center"><div align="center">VT<br>Pendente</div></td>
    <td width="10%" align="center"><div align="center">Tx. Sucesso VT<br>Pendente</div></td>
      </tr>
  <!--#include file="AbreConexao.asp"-->
  <%

SQL = SQL & " SELECT NomeAbreviado, "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVSucesso' THEN Qtde ELSE 0 END) as SrvSucesso, "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVInsucesso' THEN Qtde ELSE 0 END) as SrvInSucesso, "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVPendente' THEN Qtde ELSE 0 END) as SrvPendente, "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTSucesso' THEN Qtde ELSE 0 END) as VTSucesso, "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTInsucesso' THEN Qtde ELSE 0 END) as VTInSucesso, "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTPendente' THEN Qtde ELSE 0 END) as VTPendente, "

SQL = SQL & " CASE (convert(decimal(20,8),ISNULL((  "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVSucesso' THEN Qtde ELSE 0 END)+ "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVInsucesso' THEN Qtde ELSE 0 END)+ "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVPendente' THEN Qtde ELSE 0 END) "
SQL = SQL & " ),0))) WHEN 0 THEN 0 ELSE ISNULL(round((convert(decimal(20,8),ISNULL(  "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVSucesso' THEN Qtde ELSE 0 END) "
SQL = SQL & " ,0))) / (convert(decimal(20,8),ISNULL(  "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVSucesso' THEN Qtde ELSE 0 END)+ "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVInsucesso' THEN Qtde ELSE 0 END)+ "
SQL = SQL & " SUM(CASE WHEN Resultado = 'SRVPendente' THEN Qtde ELSE 0 END) "
SQL = SQL & " ,0))) * 100,2),0) END AS Tx_SucessoSRV, "

SQL = SQL & " CASE (convert(decimal(20,8),ISNULL((  "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTSucesso' THEN Qtde ELSE 0 END)+ "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTInsucesso' THEN Qtde ELSE 0 END)+ "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTPendente' THEN Qtde ELSE 0 END) "
SQL = SQL & " ),0))) WHEN 0 THEN 0 ELSE ISNULL(round((convert(decimal(20,8),ISNULL(  "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTSucesso' THEN Qtde ELSE 0 END) "
SQL = SQL & " ,0))) / (convert(decimal(20,8),ISNULL(  "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTSucesso' THEN Qtde ELSE 0 END)+ "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTInsucesso' THEN Qtde ELSE 0 END)+ "
SQL = SQL & " SUM(CASE WHEN Resultado = 'VTPendente' THEN Qtde ELSE 0 END) "
SQL = SQL & " ,0))) * 100,2),0) END AS Tx_SucessoVT "

SQL = SQL & " FROM ( "
SQL = SQL & " SELECT NomeAbreviado,    'SRV'+isnull(IDP.Status,'Pendente') AS Resultado, COUNT(*) AS Qtde "
SQL = SQL & " FROM          "
SQL = SQL & " NET_IMPERIUM AS I LEFT OUTER JOIN "
SQL = SQL & " NET_IMPERIUMDeParaBaixa AS IDP ON I.CodigoBaixa = IDP.Codigo "
SQL = SQL & " WHERE      "
SQL = SQL & " DAY(I.Data) = DAY(GETDATE()) AND "
SQL = SQL & " MONTH(I.Data) = MONTH(GETDATE()) AND  "
SQL = SQL & " YEAR(I.Data) = YEAR(GETDATE()) AND Turno NOT LIKE '%VT%' "
SQL = SQL & " GROUP BY IDP.Status, NomeAbreviado "
SQL = SQL & " UNION ALL "
SQL = SQL & " SELECT  NomeAbreviado,   'VT'+isnull(IDP.Status,'Pendente') AS Resultado, COUNT(*) AS Qtde "
SQL = SQL & " FROM          "
SQL = SQL & " NET_IMPERIUM AS I LEFT OUTER JOIN "
SQL = SQL & " NET_IMPERIUMDeParaBaixa AS IDP ON I.CodigoBaixa = IDP.Codigo "
SQL = SQL & " WHERE      "
SQL = SQL & " DAY(I.Data) = DAY(GETDATE()) AND "
SQL = SQL & " MONTH(I.Data) = MONTH(GETDATE()) AND  "
SQL = SQL & " YEAR(I.Data) = YEAR(GETDATE()) AND "
SQL = SQL & " Turno LIKE '%VT%' "
SQL = SQL & " GROUP BY IDP.Status, NomeAbreviado ) AS CONSOLIDADO "
SQL = SQL & " GROUP BY NomeAbreviado ORDER BY Tx_SucessoSRV DESC, Tx_SucessoVT DESC"

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

SrvSucesso = RSBUSCAS("SrvSucesso")
SrvSucesso = FormatNumber(SrvSucesso,0,-1,0,-2)

SrvInsucesso = RSBUSCAS("SrvInsucesso")
SrvInsucesso = FormatNumber(SrvInsucesso,0,-1,0,-2)

SrvPendente = RSBUSCAS("SrvPendente")
SrvPendente = FormatNumber(SrvPendente,0,-1,0,-2)

Tx_SucessoSRV = RSBUSCAS("Tx_SucessoSRV")
Tx_SucessoSRV = FormatNumber(Tx_SucessoSRV,0,-1,0,-2)

VTSucesso = RSBUSCAS("VTSucesso")
VTSucesso = FormatNumber(VTSucesso,0,-1,0,-2)

VTInsucesso = RSBUSCAS("VTInsucesso")
VTInsucesso = FormatNumber(VTInsucesso,0,-1,0,-2)

VTPendente = RSBUSCAS("VTPendente")
VTPendente = FormatNumber(VTPendente,0,-1,0,-2)

Tx_SucessoVT = RSBUSCAS("Tx_SucessoVT")
Tx_SucessoVT = FormatNumber(Tx_SucessoVT,0,-1,0,-2)

	If Chamadas > 0 Then
				Tx_Conversao = (CDbl(Pedido))/ (CDbl(Chamadas) )*100
	Else
	Tx_Conversao = 0
	End If
	Tx_Conversao = FormatNumber(Tx_Conversao,2,-1,0,-2)

	If Chamadas > 0 Then
				Tx_ConversaoOutros = (CDbl(Pedido)+CDbl(Producao)+CDbl(Embalado)+CDbl(Entrega)+CDbl(Entregue))/ (CDbl(Chamadas) )*100
	Else
	Tx_ConversaoOutros = 0
	End If
	Tx_ConversaoOutros = FormatNumber(Tx_ConversaoOutros,2,-1,0,-2)

	If Chamadas > 0 Then
				Tx_Marcacao = (CDbl(Marcacoes))/ (CDbl(Chamadas) )*100
	Else
	Tx_Marcacao = 0
	End If
	Tx_Marcacao = FormatNumber(Tx_Marcacao,2,-1,0,-2)

	SrvSucesso_TOTAL = SrvSucesso_TOTAL + CDbl(SrvSucesso)
	SrvInsucesso_TOTAL = SrvInsucesso_TOTAL + CDbl(SrvInsucesso)
	SrvPendente_TOTAL = SrvPendente_TOTAL + CDbl(SrvPendente)

	VTSucesso_TOTAL = VTSucesso_TOTAL + CDbl(VTSucesso)
	VTInsucesso_TOTAL = VTInsucesso_TOTAL + CDbl(VTInsucesso)
	VTPendente_TOTAL = VTPendente_TOTAL + CDbl(VTPendente)


	If Chamadas_TOTAL > 0 Then
				Tx_Conversao_Total = (CDbl(Pedido_TOTAL))/ (CDbl(Chamadas_TOTAL) )*100
	Else
	Tx_Conversao_Total = 0
	End If
	Tx_Conversao_Total = FormatNumber(Tx_Conversao_Total,2,-1,0,-2)

	If Chamadas_TOTAL > 0 Then
				Tx_ConversaoOutros_Total = (CDbl(Pedido_TOTAL)+CDbl(Producao_TOTAL)+CDbl(Embalado_TOTAL)+CDbl(Entrega_TOTAL)+CDbl(Entregue_TOTAL))/ (CDbl(Chamadas_TOTAL) )*100
	Else
	Tx_ConversaoOutros_Total = 0
	End If
	Tx_ConversaoOutros_Total = FormatNumber(Tx_ConversaoOutros_Total,2,-1,0,-2)


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

If Tx_SucessoSRV > 60 Then
IndConversaoSRV = "BallGreen"
End If
If Tx_SucessoSRV = 60 Then
IndConversaoSRV = "BallYellow"
End If
If Tx_SucessoSRV < 60 Then
IndConversaoSRV = "BallRed"
End If

If Tx_SucessoVT > 60 Then
IndConversaoVT = "BallGreen"
End If
If Tx_SucessoVT = 60 Then
IndConversaoVT = "BallYellow"
End If
If Tx_SucessoVT < 60 Then
IndConversaoVT = "BallRed"
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
    <td width="10%" align="left"><%=RSBuscas("NomeAbreviado")%></td>
    <td width="10%" align="left"><div align="center"><%=SrvSucesso%></td>    
    <td width="10%" align="center"><div align="center"><%=SrvInsucesso%></div></td>
    <td width="10%" align="center"><div align="center"><%=SrvPendente%></div></td>
    <td width="10%" align="center"><div align="center"><%=Tx_SucessoSRV%>% <img src="../imagens/<%=IndConversaoSRV%>.gif" width="15" height="18" border="0"></div></td>
    <td width="10%" align="center"><div align="center"><%=VTSucesso%></div></td>
    <td width="10%" align="center"><div align="center"><%=VTInsucesso%></div></td>
    <td width="10%" align="center"><div align="center"><%=VTPendente%>  </div></td>
    <td width="10%" align="center"><div align="center"><%=Tx_SucessoVT%>% <img src="../imagens/<%=IndConversaoVT%>.gif" width="15" height="18" border="0"></div></td>
  </tr>
  <%
	RSBUSCAS.Movenext
Loop
%>
<tr bgcolor="#cdd5da">
    <td colspan="3" align="left"><div align="left"><b>Total</b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=SrvSucesso_Total%></b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=SrvInsucesso_Total%></b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=SrvPendente_Total%></b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=Tx_SucessoSRV_Total%>%</b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=VTSucesso_Total%></b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=VTInsucesso_Total%></b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=VTPendente_Total%></b></div></td>
    <td width="10%" align="center"><div align="center"><b><%=Tx_SucessoVT_Total%>%</b></div></td>

  </tr>
</table>
    <input type="hidden" name="txtAcumulado" value="<%=Acumulado%>">
    <input type="hidden" name="txtDiario" value="<Diario%>">




</body>
</html>



