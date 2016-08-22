
<%
		Ultimo = Year(Now)&"-"&Month(Now)&"-"&Day(Now)&" "&TimeValue(Now)
		IP = request.servervariables("REMOTE_ADDR")
		Usuario = Session("usuario")

SQL = "INSERT INTO dbo.SIS_VolumeAcessos (LOGIN, DATA, IP) VALUES ('"&Usuario&"', CONVERT(DATETIME, '"&Ultimo&"', 102), '" & IP & "') "
	
		Set Tabela = Conexao.Execute(SQL)

		Response.Redirect "entrada.asp"
%>