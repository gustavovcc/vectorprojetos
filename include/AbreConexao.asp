<%

	Set Conexao = server.createobject("ADODB.Connection")
	Conexao.Provider = "SQLOLEDB"
	Conexao.Properties("Data Source").Value = "localhost"
	Conexao.Properties("Initial Catalog").Value = "Planejamento"
	Conexao.Properties("User ID").Value = "planejamento"
	Conexao.Properties("Password").Value = "dna@123"
	Conexao.Open

%>
