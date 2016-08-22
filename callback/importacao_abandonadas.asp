<%
Dim Enviar, campos, contador, Tipo
Enviar = Request.QueryString("Enviar")
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>WeDo Serviços</title>
<link rel="stylesheet" href="../include/pgo.css" type="text/css">
<script language="javascript">
	function ValidaDados() {
		if (confirm("Confirma a Inclusão dos Dados?")) {
				frmDados.action = "importacao_abandonadas.asp?Enviar=S";
				frmDados.submit();
		}
		return false;
	}
</script>
</head>
<body>
<form method="POST" name="frmDados">
  <table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
    <tr bgcolor="#EBF3F1">
      <td width="100%" bgcolor="#EBF3F1"> <p align="center"><b>Importação - Abandonados - Call Back</b></td>
    </tr>
    <!--#include file="AbreConexao.asp"-->
    <tr>
      <td> <p align="left">
          <input name="Importar" type="submit" onClick="return ValidaDados();" id="Importar" value="Importar" class="formbutton">
      </td>
    </tr>
  </table>
</form>
<%
Sub Inserir
%>
<!--#include file="AbreConexao.asp"-->
<%
SQL1 = "INSERT INTO PABX...calls ( id_campaign, phone, retries, dnc ) select '9' as id_campaign, '0' + right(callerid,10), '0' as retries, '0' as dnc from PABX...call_entry WHERE id_queue_call_entry in (1,2) and status = 'abandonada'  and callerid <> 'unknown' and callerid <> '9999' and callerid <> '88888888' and callerid <> '0008521812422' and day(datetime_entry_queue) = day(getdate()) and month(datetime_entry_queue) = month(getdate()) and year(datetime_entry_queue) = year(getdate()) and right(callerid,10) not in ( SELECT right(phone,10) FROM PABX...calls where id_campaign = '9' ) group by callerid "
Conexao.Execute(SQL1)
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Importando Clientes Abandonados');"
		Response.Write "</script>"

SQL2 = "UPDATE PABX...campaign SET ESTATUS = 'A' WHERE ID = '9' "
Conexao.Execute(SQL2)
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Ativando Campanha!');"
		Response.Write "</script>"

If Enviar <> "" Then
		If Err.number = 0 Then
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Operação realizada com sucesso!');"
		Response.Write "location.href='importacao_abandonadas.asp';"
		Response.Write "</script>"
	Else
		If Err.number <> 0 Then
			Response.Write "<script language=""JavaScript"">alert(""Ocorreu um erro de execução\n\nErro:" & Err.number & "\n" & Err.description & _
				"\n\nInforme esse erro ao desenvolvedor"");</script>"
		End If
	End If
Else
		Response.Write "<script language=""JavaScript"">alert(""Alguma informação estava vazia. Favor preencher."");</script>"
End If
Set Conexao = Nothing
End Sub
If Enviar = "S"  Then
	Inserir
End If
%>
&nbsp;
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td colspan="7"> <p align="center"><b>Status Clientes Abandonados - Call Back</b></td>
  </tr>
  <tr bgcolor="#EBF3F1">
    <td width="5%" align="center"> <p align="center"><i>Telefone</i></td>
    <td width="15%" align="center"><div align="center">Data Recep.</div></td>
    <td width="15%" align="center"><div align="center">Data Ativo</div></td>
    <td width="15%" align="center"><div align="center">Status Ativo</div></td>
    <td width="15%" align="center"><div align="center">Tentativas</div></td>
      </tr>
<!--#include file="AbreConexao.asp"-->
<%

SQL = " SELECT     RIGHT(CE.callerid, 10) AS callerid, CE.status AS StatusReceptivo, C.status AS StatusAtivo, CE.datetime_entry_queue AS DataReceptivo,  "
SQL = SQL & "                       C.datetime_originate AS DataAtivo, C.retries "
SQL = SQL & " FROM         PABX...call_entry AS CE LEFT OUTER JOIN "
SQL = SQL & "                       PABX...calls AS C ON RIGHT(CE.callerid, 10) = RIGHT(C.phone, 10) "
SQL = SQL & " WHERE     (DAY(CE.datetime_entry_queue) = DAY(GETDATE())) AND (MONTH(CE.datetime_entry_queue) = MONTH(GETDATE())) AND (YEAR(CE.datetime_entry_queue)  "
SQL = SQL & "                       = YEAR(GETDATE())) AND (CE.status = 'abandonada') AND (CE.callerid <> '9999' and callerid <> '88888888' and CE.callerid <> 'unknown') AND (CE.id_queue_call_entry in (1,2)) "
SQL = SQL & " ORDER BY StatusAtivo, DataAtivo, DataReceptivo "

Set RSBUSCAS = server.createobject("ADODB.Recordset")
'RESPONSE.WRITE SQL
RSBUSCAS.Open SQL, Conexao
i = 0
Do While Not RSBUSCAS.EOF
%>
  <tr>
    <td width="5%" align="center"><%=RSBuscas("callerid")%></td>
    <td width="15%" align="center"><%=RSBuscas("DataReceptivo")%></td>
    <td width="15%" align="center"><%=RSBuscas("DataAtivo")%></td>
    <td width="15%" align="center"><%=RSBuscas("StatusAtivo")%></td>
    <td width="15%" align="center"><%=RSBuscas("retries")%></td>
  </tr>
    <%
  Response.Flush
  i=i+1
	RSBUSCAS.Movenext
Loop
%>
</table>
</body>
</html>