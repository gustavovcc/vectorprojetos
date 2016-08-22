<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<style type="text/css">
/*Configuração do CSS do Letreiro*/
<!--
a {font-family: Verdana; font-size: 11px; }
a:link {text-decoration: underline; color: #000000}
a:active {text-decoration: underline; color: #000000}
a:visited {text-decoration: underline; color: #000000}
-->

#pscroller1{
width: 357px;
height: 40px;
border: 0px;
padding: 10px;
background-color: white;
}
</style>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Documento sem título</title>
<link rel="stylesheet" href="../include/pgo.css" type="text/css">
	<link rel="StyleSheet" href="dtree.css" type="text/css" />
	<script type="text/javascript" src="dtree.js"></script>
</head>



<body>
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0">
  <tr>
    <td width="15%" rowspan="2" valign="top"><%
'//Acesso Administrador
Acesso = Session("Acesso")
If Acesso = "Administrador" or Acesso = "Supervisor" Then
Hoje = Month(Now)&"-"&Day(Now)&"-"&Year(Now)
%>
<div class="dtree">

	<p><a href="javascript: d.openAll();">Abrir Todos</a> | <a href="javascript: d.closeAll();">Fechar Todos</a></p>

	<script type="text/javascript">
		<!--

		d = new dTree('d');

		d.add(0,-1,'Grupo Vector', 'http://www.vectorcontactcenter.com.br');
		d.add(1,0,'Administração','','Administração dos Sistemas','','','img/imgfolder.gif');
		d.add(2,1,'Cadastro Usuários','../usuarios.asp','Cadastrar Novo Usuário');
		d.add(3,0,'Links Úteis');
		<!--d.add(33,3,'Pizza Hut','www.pizzahut-ce.com.br','', '_blank');
		d.add(34,3,'Call Flex','http://10.10.4.32/manager/','', '_blank');
		d.add(35,3,'Atom','http://10.10.4.6/Atom','', '_blank');
		d.add(36,3,'Correios - Busca de CEP','http://www.buscacep.correios.com.br','', '_blank');
		d.add(39,3,'Consulta Portabilidade','http://consultanumero.abrtelecom.com.br/consultanumero/consulta/consultaSituacaoAtualCtg','', '_blank');
		d.add(40,3,'vDesk','http://10.10.4.60/','', '');
		d.add(20,0,'Relatórios');
		d.add(21,20,'NET Vendas');
		d.add(51,21,'Ranking de Vendas - Por Vendedor', '../relatorios/tempoRealPedidos_Comissao.asp');
		
		d.add(22,20,'NET Agendamento');
		d.add(61,22,'Ranking de Agendamento - Por Atendente', '../relatorios/report_Agendamento_PorAtendente.asp');
		d.add(62,22,'Painel Agendamento - TV', '../relatorios/tempoRealAgendamento_Painel.asp');
		
		d.add(23,20,'NET Instalação');
		d.add(71,23,'Ranking de Instalação - Por Instalador', '../relatorios/tempoRealPedidos_Comissao.asp');
		
		d.add(24,20,'Outras');
		d.add(71,21,'Ranking Outras - Por Vendedor', '../relatorios/tempoRealPedidos_Comissao-.asp');
		
		d.add(99,0,'Links Antigos','','','','img/trash.gif');
		document.write(d);

		//-->
	</script>
<%
End if
%>

</td>

  <td width="33%" rowspan="2" align="center" valign="top"><table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td colspan="9"> <p align="center"><b>Painel de Resultados até <%=now%></b></td>
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
SQL = SQL & " SUM(Improdutiva) as Improdutiva, "
SQL = SQL & " SUM(Atendida) as Atendidas, "
SQL = SQL & " SUM(Telefonia) as Telefonia, "
SQL = SQL & " SUM(Agendamento)+ "
SQL = SQL & " SUM(Atendida)+ "
SQL = SQL & " SUM(Telefonia) as Total, "
SQL = SQL & " CASE (convert(decimal(20,8),ISNULL((  "
SQL = SQL & " ISNULL(SUM(Atendida),0) "
SQL = SQL & " ),0)))  "
SQL = SQL & " WHEN 0 THEN 0 ELSE ISNULL(round((convert(decimal(20,8),ISNULL(  "
SQL = SQL & " ISNULL(SUM(Agendamento) "
SQL = SQL & " ,0)  "
SQL = SQL & " ,0))) / (convert(decimal(20,8),ISNULL(  "
SQL = SQL & " ISNULL(SUM(Atendida),0) "
SQL = SQL & " ,0))) * 100,2),0) END AS TxAgendamento, "

SQL = SQL & " CASE (convert(decimal(20,8),ISNULL((  "
SQL = SQL & " ISNULL(SUM(Atendida),0) "
SQL = SQL & " ),0)))  "
SQL = SQL & " WHEN 0 THEN 0 ELSE ISNULL(round((convert(decimal(20,8),ISNULL(  "
SQL = SQL & " ISNULL(SUM(Improdutiva) "
SQL = SQL & " ,0)  "
SQL = SQL & " ,0))) / (convert(decimal(20,8),ISNULL(  "
SQL = SQL & " ISNULL(SUM(Atendida),0) "
SQL = SQL & " ,0))) * 100,2),0) END AS TxImprod, "

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
SQL = SQL & " SUM(cast(isnull(SRAG.Improdutiva,0) as int)) as Improdutiva,  "
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

TxImprod = RSBUSCAS("TxImprod")
TxImprod = FormatNumber(TxImprod,2,-1,0,-2)

TxAtend = RSBUSCAS("TxAtend")
TxAtend = FormatNumber(TxAtend,2,-1,0,-2)

%>

  <tr>
	<td width="25%" align="center"><p align="left" >NET Agendamento</div></td>
	<td width="25%" align="center"><p align="center" > % Agend.: <%=TxAgendamento%>%</div></td>
    <td width="25%" align="center"><p align="center" > % Impr.: <%=TxImprod%>%</div></td>
	<td width="25%" align="center"><p align="center" > % Atend.: <%=TxAtend%>%</div></td>
   </tr>

<%
  Response.Flush
  i=i+1
	RSBUSCAS.Movenext
Loop
%>
</table>

</td>
  </tr>
</table>



</div>

</body>
</html>
