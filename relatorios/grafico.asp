<% Option Explicit %>

<!--#include virtual="adovbs.inc"-->

<%
'definindo a largura e a altura em pixels 
Const grafaltura = 300
Const graflargura = 600
Const barImage = "img_graf.gif"

sub BarChart(data, rotulos , titulo, eixos)
   'Imprime o cabeçalho
   Response.Write("<TABLE CELLSPACING=0 CELLPADDING=1 BORDER=0 WIDTH=" & graflargura & ">" 
& chr(13))
   Response.Write("<TR><TH COLSPAN=" & UBound(data) - LBound(data) + 2 & ">")
   Response.Write("<FONT SIZE=+2>" & titulo & "</FONT></TH></TR>" & chr(13))
   Response.Write("<TR><TD VALIGN=TOP ALIGN=RIGHT>" & chr(13))

   'encontra o maior valor
   Dim maior_valor
   maior_valor = data(LBound(data))

   Dim i
   for i = LBound(data) to UBound(data) - 1
  	if data(i) > maior_valor then maior_valor = data(i)
   next

   'imprime o maior valor no topo do gráfico
   Response.Write("<b>" & maior_valor & "</b>-" & "</TD>")

   Dim largura_percentual
   largura_percentual = CInt((1 / (UBound(data) - LBound(data) + 1)) * 100)

For i = LBound(data) to UBound(data) - 1
  Response.Write(" <TD VALIGN=BOTTOM ROWSPAN=2 WIDTH=" & largura_percentual & "% >" & chr(13))
  Response.Write("   <IMG SRC=""" & barImage & """ WIDTH=100% HEIGHT=" & CInt(data(i)/maior_valor
 * grafaltura) & ">" & chr(13))
  Response.Write(" </TD>" & chr(13))
Next

  Response.Write("</TR>")
  Response.Write("<TR><TD VALIGN=BOTTOM ALIGN=RIGHT><b>0</b></TD></TR>")

  'Imprime o rodape
  Response.Write("<TR><TD ALIGN=RIGHT VALIGN=BOTTOM>" & eixos & "</TD>" & chr(13))
  for i = LBound(rotulos) to UBound(rotulos) - 1
    Response.Write("<TD VALIGN=BOTTOM ALIGN=CENTER>" & rotulos(i) & "</TD>" & chr(13))
  next
  Response.Write("</TR>" & chr(13))
  Response.Write("</TABLE>")
end sub

'abre conexao com banco de dados
Dim objConnection
Set objConnection = Server.CreateObject("ADODB.Connection")
objConnection.Open "DSN=Faltas"

Dim SQL
SQL = "SELECT Aluno,Faltas FROM Alunos"

Dim rsFaltas
Set rsFaltas = Server.CreateObject("ADODB.Recordset")
rsFaltas.Open SQL, objConnection, adOpenStatic
Encontra o total de registros do arquivo
Dim numRegistros
numRegistros = rsFaltas.RecordCount
Define os vetores que irão armazenar as faltas e o nome dos alunos
Dim VetorFaltas(), VetorNomes()
Redim VetorFaltas(numRegistros)
Redim VetorNomes(numRegistros)

Dim i
for i = 0 to numRegistros-1
	VetorFaltas(i) = rsFaltas("Faltas")
	VetorNomes(i) = rsFaltas("Aluno")
	rsFaltas.MoveNext
next

%>

<HTML>
<BODY>
<CENTER>
<% BarChart VetorFaltas,VetorNomes,"Faltas dos Alunos da 1a. Serie - Março/Abril ","Alunos" %>
</CENTER>
</BODY>
</HTML>

<%
	rsFaltas.Close
	Set rsFaltas = Nothing

	objConnection.Close
	Set objConnection = Nothing
%>