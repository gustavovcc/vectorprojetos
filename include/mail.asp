<HTML>

<HEAD>
<TITLE></TITLE>
</HEAD>

<BODY>

  <FORM ACTION="http://form.ultramail.com.br/" METHOD="POST">
  <P>
<!--
  Formul�rio do cliente.
  Especifique abaixo os campos que deseja enviar para e-mail.
  Caso o campo assunto n�o seja preenchido, o sistema ir� enviar o e-mail com o assunto Formul�rio UltraMail
-->
Nome: <BR><INPUT TYPE="text" NAME="nome" SIZE="24"><BR>
E-Mail: <BR><INPUT TYPE="text" NAME="email" SIZE="24"><BR>
Assunto: <BR><INPUT TYPE="text" NAME="assunto" SIZE="24"><BR>
Mensagem: <BR><TEXTAREA NAME="mensagem" ROWS="8" COLS="20"></TEXTAREA>

<!--
  Chave de autentica��o no UltraMail para o MailBox.
  Se a senha do MailBox for alterada esta chave dever� ser gerada novamente atrav�s do seu painel de controle.
-->
    <INPUT TYPE="hidden" NAME="key" VALUE="eJwBxAA7/6mPHJg0fcPlUtRzWobwgg34CS35iAoVTDBHZ6mNvwUQRm9ybVVsdHJhTWFpbBb7zRmOE3EfIZe4qvNnLYM7altCfvUqLPKSYgAPzonA/3qaK5uw66adj9uCplnlsNZzTreX9O0Ot+AMge9lZQ4rhd7rhmeIA/K2FdLbohtqj3d93OtHjsk0Y74w5YTIGAkzysgQ7W6VSHoN8YiTZOYEsr0gZE+jHLWPOgI9HRSeSIhvvOVONxQeOjMUYUplOklxAb2A83MBaFw0">

<!--
  Pagina de conclus�o do formul�rio de envio. Altere para a p�gina desejada
-->
    <INPUT TYPE="hidden" NAME="redirect" VALUE="http://wedoservicos.com.br/PaginaDeResposta.html">

    <INPUT TYPE="submit" VALUE="Enviar">
    <INPUT TYPE="reset" VALUE="Limpar">
  </P>
  </FORM>

</BODY>
</HTML>