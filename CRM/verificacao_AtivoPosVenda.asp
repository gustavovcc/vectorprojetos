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
<title>Administra&ccedil;&atilde;o de Usu�rio</title>
<link rel="stylesheet" href="../include/pgo.css" type="text/css">
<script language="javascript" src="../Include/MostraCalendario.js"></script>
<script language="javascript">
	function ValidaDados() {
		if (confirm("Confirma a busca?")) {
				frmDados.action = "verificacao_AtivoPosVenda.asp?Enviar=S";
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
      <td colspan="2" bgcolor="#EBF3F1"> <p align="center"><b>Localiza��o Clientes - P�s-Venda</b></td>
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
    <td colspan="16"> <p align="center"><b>Rela��o Clientes</b></td>
  </tr>
  <tr bgcolor="#EBF3F1">
    <td width="2%" align="center"><p align="center"><i>Ficha</i></td>
    <td width="2%" align="center"><p align="center"><i>Hist&oacute;rico</i></td>
    <td width="2%" align="center"><p align="center"><i>Telefone</i></td>
    <td width="20%" align="center"><p align="center"><i>Nome</i></td>
    <td width="2%" align="center"><p align="center"><i>Valor</i></td>
    <td width="10%" align="center"><p align="center"><i>Data</i></td>
    <td width="5%" align="center"><p align="center"><i>Filial</i></td>
    <td width="5%" align="center"><p align="center"><i>Loja</i></td>
   </tr>

<!--#include file="AbreConexao.asp"-->
<%
Data = Year(Data)&"-"&Month(Data)&"-"&Day(Data)
If Data = "" Then
	Data = Year(now)&"-"&Month(now)&"-"&Day(now)
End If

SQL = " SELECT * FROM ( "
SQL = SQL & " SELECT VEN_NRTELE AS TELEFONE, VEN_NMCLIE AS NOMECLIENTE, VEN_DHRECE AS DATA, VEN_VLRNOT AS VALOR, "
SQL = SQL & " CASE WHEN VEN_NRLOJA = '3' THEN 'O. Paiva' ELSE  "
SQL = SQL & " CASE WHEN VEN_NRLOJA = '5' THEN 'Maraponga' ELSE  "
SQL = SQL & " CASE WHEN VEN_NRLOJA = '4' THEN '' ELSE 'Outros' END END END as FILIAL, 'Pastel na Hora' as LOJA "
SQL = SQL & " FROM PastelNAHORA_DELIVERY..TB_MOVVENDA "
SQL = SQL & " WHERE VEN_TPVEND = 2 AND VEN_STATUS = 0 and  "
SQL = SQL & " day(VEN_DHRECE) = day(getdate()-1) and  "
SQL = SQL & " Month(VEN_DHRECE) = Month(getdate()) and  "
SQL = SQL & " Year(VEN_DHRECE) = Year(getdate()) and "
SQL = SQL & " VEN_NRTELE NOT IN (SELECT DDDTELEFONE FROM DNA_PastelnaHora..tbFichasAtendimento_PosVenda where day(DataCriacao) = day(getdate()) and Month(DataCriacao) = Month(getdate()) and Year(DataCriacao) = Year(getdate())) "
SQL = SQL & " ) A "
SQL = SQL & " ORDER BY TELEFONE, DATA "


Set RSBUSCAS = server.createobject("ADODB.Recordset")
'RESPONSE.WRITE SQL
RSBUSCAS.Open SQL, Conexao
i = 0
Do While Not RSBUSCAS.EOF
%>
  <tr>
    <td width="2%" align="center"><p align="center">
<% If RSBuscas("Telefone") <> "" Then %>    
      <a href="incluir_FichaAtendimento_PosVenda.asp?User=<%=User%>&ID=<%=RSBuscas("Telefone")%>&NomeCliente=<%=RSBuscas("NomeCliente")%>&Loja=<%=RSBuscas("Loja")%>">
        <img src="../imagens/mais.jpg" width="16" height="16" border="0">
        </a>   
<% Else %>
      <a href="cadProspects.asp?User=<%=User%>&DDDTelefone=<%=RSBuscas("Telefone")%>">
<img src="../imagens/edit.png" width="16" height="16" border="0">
<% End If%>        
    <td width="2%" align="center"><p align="center">
<% If RSBuscas("Telefone") <> "" Then %>    
<a href="historico_fichas.asp?User=<%=User%>&Telefone=<%=RSBuscas("Telefone")%>">
   <img src="../imagens/viewmag.png" width="16" height="16" border="0">
   </a>
<% Else %>
<img src="../imagens/viewmag.png" width="16" height="16" border="0">
<% End If%>
     <td width="2%" align="center"><p align="center"><%=RSBuscas("Telefone")%></td>
    <td width="20%" align="center"><p align="center"><%=RSBuscas("NomeCliente")%></td>
    <td width="2%" align="center"><p align="center">R$ <%=RSBuscas("Valor")%></td>
    <td width="10%" align="center"><p align="center"><%=RSBuscas("Data")%></td>
    <td width="5%" align="center"><p align="center"><%=RSBuscas("Filial")%></td>
    <td width="5%" align="center"><p align="center"><%=RSBuscas("Loja")%></td>
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