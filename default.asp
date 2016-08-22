<%
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma", "No-Cache"
Response.Expires = -1000
Response.Buffer = True
%>
<html>
<head>
<title>DNA - Planejamento Vector</title>
<script language="javascript">
	function ValidarUsuario() {
		if (frmDados.Usuario.value == "")	{
			alert("O campo usuário não foi preenchido! Por favor, preencha o campo usuário.");
			frmDados.Usuario.focus();
			return false;
		}
		if (frmDados.Senha.value == "") {
			alert("O campo senha não foi preenchido! Por favor, preencha o campo senha.");
			frmDados.Senha.focus();
			return false;
		}
		frmDados.submit();
	}
</script>
<link rel="stylesheet" type="text/css" href="include/estilo.css">
</head>

<body onLoad="javascript:document.frmDados.Usuario.focus();">
<br>
<form action="default.asp" method="post" name="frmDados">
<input type="hidden" name="Entrar" value="S">
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFFFFF">
          <td width="17%" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="57%" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="26%" bgcolor="#FFFFFF"><div align="right">
			<img border="0" src="imagens/vector_Logo.png" width="168" height="100"></div></td>
        </tr>
        <tr bgcolor="#FFFFFF">
          <td>&nbsp;</td>
          <td colspan="2">&nbsp;</td>
        </tr>
        <tr bgcolor="#FF9900">
          <td bgcolor="#0080FF">&nbsp;</td>
          <td colspan="2" bgcolor="#0080FF">&nbsp;</td>
        </tr>
        <tr bgcolor="#FF9900">
          <td bgcolor="#0080FF">&nbsp;</td>
          <td colspan="2" bgcolor="#0080FF">&nbsp;</td>
        </tr>
        <tr bgcolor="#FF9900">
          <td colspan="3" bgcolor="#81BEF7"><table width="18%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="33%" bgcolor="#81BEF7" class="textos"><span class="style24">Usu&aacute;rio</span></td>
              <td width="67%" bgcolor="#81BEF7"><input name="Usuario" type="text" class="textoform" id="Usuario"></td>
              <td width="67%" rowspan="3" bgcolor="#81BEF7">&nbsp;  </td>
              <td width="67%" rowspan="3" bgcolor="#81BEF7">&nbsp;</td>
            </tr>
            <tr>
              <td bgcolor="#81BEF7" class="textos"><span class="style24">Senha</span></td>
              <td bgcolor="#81BEF7"><input name="Senha" type="password" class="textoform" id="Senha"></td>
              </tr>
            <tr>
              <td bgcolor="#81BEF7">&nbsp;</td>
              <td bgcolor="#81BEF7"><input name="btEntrar" onClick="return ValidarUsuario();" type="submit" class="textoform" id="btEntrar" value="Entrar"></td>
              </tr>
          </table></td>
        </tr>
        <tr bgcolor="#FF9900">
          <td bgcolor="#0080FF">&nbsp;</td>
          <td colspan="2" bgcolor="#0080FF">&nbsp;</td>
        </tr>
        <tr bgcolor="#FF9900">
          <td bgcolor="#0080FF">&nbsp;</td>
          <td colspan="2" bgcolor="#0080FF">&nbsp;</td>
        </tr>
      </table>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>
<%
Sub CheckLogin

Dim Conexao, sql, RS, username, userpwd, usercargo, Matricula, recLogin

username = Request.Form("Usuario")
userpwd = Request.Form("Senha")
%>

<!--#include file="include/AbreConexao.asp"-->
<%

sql = "select * from SIS_usuarios where Status = 1 and (usuario = '" & LCase(username) & "' "
sql = sql & ") and (senha = '" & Trim(userpwd) & "')"


Set Tabela = Conexao.Execute(sql)

If Tabela.BOF And Tabela.EOF Then

	Response.Write "<script language='javascript'>"
	Response.Write "alert('Não foi possível logar. Tente novamente.');"
	Response.Write "</script>"
	Response.Redirect "default.asp"
	Response.End

Else

  Session("UserLoggedIn") = "true"
	Session("Nome")= Tabela("Nome")
	Session("usuario")= Tabela("usuario")
	Session("acesso")= Tabela("acesso")
    Session.TimeOut = 360
    Response.Write "alert('Usuário Logado!');"

		z = Trim(Tabela.Fields("Acessos").value)
		Ultimo = Year(Now)&"-"&Month(Now)&"-"&Day(Now)&" "&TimeValue(Now)
		If isnull(z) Then
			z = 0
		End If
		z = CDBl(z) + 1
		SQL2 = "UPDATE SIS_Usuarios SET ACESSOS = " & z & " WHERE (usuario = '" & LCase(username) & "') "
		Set Tabela2 = Conexao.Execute(SQL2)
		SQL3 = "UPDATE SIS_Usuarios SET UltimoAcesso = CONVERT(DATETIME, '"&Ultimo&"', 102) WHERE (usuario = '" & LCase(username) & "') "
		Set Tabela3 = Conexao.Execute(SQL3)

If Trim(userpwd) = "1" or Trim(userpwd) = "123" then
		Response.Redirect "altera.asp?MSG=Sim"
else
		Response.Redirect "entrada.asp"
End If


End If

recLogin.close
conexao.close
Set recLogin = Nothing
Set SQL = Nothing
Set Conexao = Nothing

End Sub

If Request.Form("Entrar") = "S" Then
	CheckLogin
End If
%>
</form>
</body>
</html>
