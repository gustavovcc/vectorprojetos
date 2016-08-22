<%
Dim i, Data, Enviar, ID
Data = Request.QueryString("Dia")
Enviar = Request.QueryString("Enviar")
Faixa = Request.QueryString("Faixa")
ID = Request.QueryString("ID")
Ordem = Request.QueryString("Ordem")
User = Trim(Session("usuario"))
Empresa = Trim(Session("Empresa"))
IDConsultor = Trim(Session("IDConsultor"))
Contrato = Trim(Session("Contrato"))

If User = "" Then
	Response.Write "<script language='javascript'>"
	Response.Write "alert('Efetue o logon novamente.');"
	Response.Write "parent.parent.top.location.href='../default.asp';"
	Response.Write "</script>"
	Response.End
End If
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Administra&ccedil;&atilde;o de Usuário</title>
<link rel="stylesheet" href="../include/pgo.css" type="text/css">
<script language="javascript" src="../Include/MostraCalendario.js"></script>
<script language="javascript">
	function ValidaDados() {
		if (confirm("Confirma a busca?")) {
				frmDados.action = "verificacao_prospects.asp?Enviar=S";
				frmDados.submit();
		}
		return false;
	}
</script>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="javascript:document.frmDados.txtDDDTelefone.focus();">
<form method="POST" name="frmDados">
    <input type="hidden" name="txtUsuario" value="<%=User%>">
  <table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center" >
    <tr bgcolor="#EBF3F1">
      <td colspan="2" bgcolor="#EBF3F1"> <p align="center"><b>Localização Clientes</b></td>
    </tr>
    <tr>
      <td width="10%"><p align="center"><i>DDDTelefone</i></td>
      <td width="10%"><p align="center"><i>Nome Cliente</i></td>
    </tr>
    <tr>
      <td><p align="center"><input name="txtDDDTelefone" type="text" class="formfield" size="15" maxlength="10">
      </td>

      <td><p align="center"><input name="txtNomeCliente" type="text" class="formfield" size="20" maxlength="50" value="<%=Request("txtNomeCliente")%>"></td>
    </tr>
    <tr>
      <td colspan="2"> <p align="center">
          <input name="Buscar" type="submit" onClick="return ValidaDados();" id="Buscar" value="Buscar" class="formbutton">
      </td>
    </tr>
  </table>
</form>
<%
Sub Consultar
%>
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td colspan="16"> <p align="center"><b>Relação Clientes</b></td>
  </tr>
  <tr bgcolor="#EBF3F1">
    <td width="2%" align="center"><i>Ficha</i></td>
    <td width="2%" align="center"><i>Hist&oacute;rico</i></td>
    <td width="2%" align="center"><p align="center"><i>Mapa</i></td>
    <td width="2%" align="center"><p align="center"><i>DDD Telefone</i></td>
    <td width="20%" align="center"><p align="center"><i>Nome Cliente</i></td>
    <td width="20%" align="center"><p align="center"><i>CEP</i></td>
    <td width="5%" align="center"><p align="center"><i>Tel. Contato</i></td>
    <td width="8%" align="center"><p align="center"><i>Observa&ccedil;&atilde;o</i></td>
   </tr>
  <tr>
    <td colspan="9"><p align="center"><a href="cadProspects.asp">Cadastrar Novo Cliente</a></td>
  </tr>   
<!--#include file="AbreConexao.asp"-->
<%
Data = Year(Data)&"-"&Month(Data)&"-"&Day(Data)
If Data = "" Then
	Data = Year(now)&"-"&Month(now)&"-"&Day(now)
End If

SQL = " SELECT TOP 10 * FROM tbProspects AS P" 
SQL = SQL & " Where DDDTelefone like '%"& Request("txtDDDTelefone") &"%' "
If Request("txtNomeCliente") <> "" Then
SQL = SQL & " and NomeCliente like '%"& Request("txtNomeCliente") &"%' "
End If
SQL = SQL & " ORDER BY NomeCliente ASC"

Set RSBUSCAS = server.createobject("ADODB.Recordset")
'RESPONSE.WRITE SQL
RSBUSCAS.Open SQL, Conexao
i = 0
Do While Not RSBUSCAS.EOF
%>
  <tr>
    <td width="2%" align="center"><p align="center">
      <a href="incluir_FichaAtendimento.asp?User=<%=User%>&ID=<%=RSBuscas("ID")%>&NomeCliente=<%=RSBuscas("NomeCliente")%>">
        <img src="../imagens/mais.jpg" width="16" height="16" border="0">
        </a>   
    <td width="2%" align="center"><p align="center">
<a href="historico_fichas.asp?User=<%=User%>&ID=<%=RSBuscas("ID")%>">
   <img src="../imagens/viewmag.png" width="16" height="16" border="0">
   </a>
     <td width="2%" align="center"><p align="center">
    <a href="mapa_prospects.asp?User=<%=User%>&ID=<%=RSBuscas("ID")%>&Endereco=<%=RSBuscas("Endereco")%> - <%=RSBuscas("Numero")%> - <%=RSBuscas("Bairro")%> - <%=RSBuscas("Cidade")%>">
    <img src="../imagens/iconemapa.png" width="16" height="16" border="0">
    </a>
    </td>
    <td width="2%" align="center"><p align="center"><%=RSBuscas("DDDTelefone")%></td>
    <td width="20%" align="center"><p align="center"><%=RSBuscas("NomeCliente")%></td>
    <td width="20%" align="center"><p align="center"><%=RSBuscas("CEP")%></td>
    <td width="8%" align="center"><p align="center"><%=RSBuscas("DDDTelefoneCEL")%></td>
    <td width="8%" align="center"><p align="center"><%=RSBuscas("Observacao")%>&nbsp;</td>
   </tr>
  <%
  Response.Flush
  i=i+1
	RSBUSCAS.Movenext
Loop
%>
  <tr>
    <td colspan="16"><p align="center">Total de <b><%=i%></b> Cliente(s) Encontrado(s) </td>
  </tr>
  <%
End Sub
If Enviar = "S"  Then
	Consultar
End If
%>
</table>
</body>
</html>