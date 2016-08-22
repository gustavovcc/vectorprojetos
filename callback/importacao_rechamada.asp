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
				frmDados.action = "importacao_rechamada.asp?Enviar=S";
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
      <td width="100%" bgcolor="#EBF3F1"> <p align="center"><b>Importação - Rechamadas - Ativo Pesquisa</b></td>
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
SQL3 = "UPDATE PABX...calls SET retries = '5' WHERE ID_Campaign = '95' "
Conexao.Execute(SQL3)
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Excluindo Clientes Antigos...');"
		Response.Write "</script>"

SQL1 = "INSERT INTO PABX...calls ( id_campaign, phone, retries, dnc ) SELECT distinct '95' as id_campaign, case len(callerid) WHEN 10 THEN '0' + right(callerid,10) ELSE + '0' + right(callerid,11) END, '0' as retries, '0' as dnc FROM tbRechamada WHERE day(Data_Rechamada) = day(getdate()-1) and month(Data_Rechamada) = month(getdate()) and year(Data_Rechamada) = year(getdate()) and left(right(callerid,11),2) in (85, 59) and Status_origem = 'terminada' and callerid <> 'unknown' "
Conexao.Execute(SQL1)
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Importando Clientes Abandonados');"
		Response.Write "</script>"

SQL2 = "UPDATE PABX...campaign SET ESTATUS = 'A' WHERE ID = '95' "
Conexao.Execute(SQL2)
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Ativando Campanha!');"
		Response.Write "</script>"

If Enviar <> "" Then
		If Err.number = 0 Then
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Operação realizada com sucesso!');"
		Response.Write "location.href='importacao_rechamada.asp';"
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
    <td width="15%" align="center"><div align="center">Origem.</div></td>
    <td width="15%" align="center"><div align="center">Rechamada</div></td>
    <td width="15%" align="center"><div align="center">Agente Origem</div></td>
    <td width="15%" align="center"><div align="center">Agente Rechamada</div></td>
    <td width="15%" align="center"><div align="center">Espera</div></td>
      </tr>
<!--#include file="AbreConexao.asp"-->
<%


SQL = " SELECT *  "
SQL = SQL & " FROM tbRechamada  "
SQL = SQL & " WHERE day(Data_Rechamada) = day(getdate()-1)  "
SQL = SQL & " and month(Data_Rechamada) = month(getdate())  "
SQL = SQL & " and year(Data_Rechamada) = year(getdate())  "
SQL = SQL & " and left(right(callerid,11),2) in (85, 59)  "
SQL = SQL & " and Status_origem = 'terminada'  "
SQL = SQL & " and callerid <> 'unknown' "

Set RSBUSCAS = server.createobject("ADODB.Recordset")
'RESPONSE.WRITE SQL
RSBUSCAS.Open SQL, Conexao
i = 0
Do While Not RSBUSCAS.EOF
%>
  <tr>
    <td width="5%" align="center"><%=RSBuscas("callerid")%></td>
    <td width="15%" align="center"><%=RSBuscas("Data_Origem")%></td>
    <td width="15%" align="center"><%=RSBuscas("Data_rechamada")%></td>
    <td width="15%" align="center"><%=RSBuscas("Agente_origem")%></td>
    <td width="15%" align="center"><%=RSBuscas("Agente_rechamada")%></td>
    <td width="15%" align="center"><%=RSBuscas("Espera")%></td>
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