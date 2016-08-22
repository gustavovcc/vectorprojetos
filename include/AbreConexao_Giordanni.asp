<%

	Set Conexao = server.createobject("ADODB.Connection")
	Conexao.Provider = "SQLOLEDB"
	Conexao.Properties("Data Source").Value = "192.168.0.195"
	Conexao.Properties("Initial Catalog").Value = "BSDigital"
	Conexao.Properties("User ID").Value = "sa"
	Conexao.Properties("Password").Value = "sa"
	Conexao.Open

%>
