<%
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma", "No-Cache"


User = Session("usuario")
If User = "" Then
	Response.Write "<script language='javascript'>"
	Response.Write "alert('Efetue o logon novamente.');"
	Response.Write "parent.parent.top.location.href='default.asp';"
	Response.Write "</script>"
	Response.End
End If

Dim Endereco
Endereco = Request.QueryString("Endereco")

%>
<html>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>BS Digital</title>
<link rel="stylesheet" href="../include/pgo.css" type="text/css">
<script type="text/javascript" src="http://maps.google.com/maps/api/js?sensor=false"></script>
<script type="text/javascript">
  var directionDisplay;
  var directionsService = new google.maps.DirectionsService();
  var map;

  function initialize() {
    directionsDisplay = new google.maps.DirectionsRenderer();
    var chicago = new google.maps.LatLng(-3.759517,-38.535008);
    var myOptions = {
      zoom:15,
      mapTypeId: google.maps.MapTypeId.ROADMAP,
      center: chicago
    }
    map = new google.maps.Map(document.getElementById("map_canvas"), myOptions);
    directionsDisplay.setMap(map);
	directionsDisplay.setPanel(document.getElementById("directionsPanel"));
  }

  function calcRoute() {
    var start = document.getElementById("start").value;
    var end = document.getElementById("end").value;
    var request = {
        origin:start,
        destination:end,
        travelMode: google.maps.DirectionsTravelMode.DRIVING
    };
    directionsService.route(request, function(response, status) {
      if (status == google.maps.DirectionsStatus.OK) {
        directionsDisplay.setDirections(response);
      }
    });
  }
</script>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="initialize();calcRoute()">
<table width="100%" border="1" bordercolorlight="#006666" bordercolordark="#ffffff" cellspacing="0" height="0" align="center">
  <tr bgcolor="#EBF3F1">
    <td colspan="2"><p align="center"><b>Rota para o Endereço</b></td>
  </tr>
  <tr bgcolor="#EBF3F1">
    <td width="50%"><p align="center"><b>Saída:</b></td>
    <td width="50%"><p align="center">Chegada:</td>
  </tr>
  <tr bgcolor="#EBF3F1">
    <td><p align="center">
<select id="start" onChange="calcRoute();">
  <option value="Av. Engenheiro Santana Júnior, 2277, Papicu, Fortaleza, CE">Pzza Hut - Casa Blanca</option>
  <option value="Av. Bezerra de Menezes, 2450, São Geraldo, Fortaleza, CE">Pzza Hut - North Shopping</option>  
  <option value="Av. Beira Mar, 2500, Meireles, Fortaleza, CE">Pizza Hut - Beira Mar</option>
  <option value="Av. Governador Flávio Ribeiro Coutinho, 115, João Pessoa, PB">Pizza Hut - João Pessoa</option>
  
</select>
</td>
    <td><p align="center">
    <select name="end" id="end" onChange="calcRoute();">
	<option value="<%=Endereco%>" selected><%=Endereco%></option>
    </select>
    </td>
  </tr>
</table>
<div id="map_canvas" style="top:10px;"></div>
<div><div id="map_canvas" style="float:left;width:70%; height:70%"></div>
<div id="directionsPanel" style="float:right;width:30%;height 30%"></div>
</body>

