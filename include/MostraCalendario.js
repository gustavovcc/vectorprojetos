function MostraCalendario(parObjeto, parSpan) {
	l_js_Objeto = parObjeto;
	l_js_Span = parSpan;
	l_js_Objeto.style.backgroundColor = "#ffff00"

	l_js_X = event.x + document.body.scrollLeft;
	if (l_js_X + 220 > document.body.offsetWidth)
		l_js_X = document.body.offsetWidth - 220;

	l_js_Y = event.y + document.body.scrollTop;
	if (l_js_Y + 140 > document.body.offsetHeight + document.body.scrollTop)
		l_js_Y = document.body.offsetHeight - 140 + document.body.scrollTop;

	if (parObjeto.value == "") { 
		l_js_Data = new Date();
		l_js_Dia = l_js_Data.getDate();
		l_js_Mes = l_js_Data.getMonth();
		l_js_Ano = l_js_Data.getYear();
	}
	else { 
		l_js_Data = parObjeto.value;
		l_js_Dia = Math.abs(l_js_Data.substring(0, 2));
		l_js_Mes = Math.abs(l_js_Data.substring(3, 5)) - 1;
		l_js_Ano = Math.abs(l_js_Data.substring(6, 10));

	}
	document.all["iframeCalendario"].style.left = l_js_X;
	document.all["iframeCalendario"].style.top = l_js_Y;
	iframeCalendario.GeraCalendario(l_js_Dia, l_js_Mes, l_js_Ano);
}

function PreenchetxtData(SpanData, txtData) {
	
	dtData = new Date();
	dtDia = dtData.getDate();
	dtMes = dtData.getMonth() + 1;
	dtAno = dtData.getYear();

	l_js_Span = SpanData;
	l_js_txtData = txtData;
		
	if (dtDia < 10) {
		dtDia = "0" + dtDia;
	}

	if (dtMes < 10) {
		dtMes = "0" + dtMes;
	}
	
	document.all[l_js_Span].style.color = "#000000";
	document.all[l_js_Span].style.fontSize = "12";
	document.all[l_js_Span].innerHTML = dtDia + "/" + dtMes + "/" + dtAno;
	document.all[l_js_txtData].value = dtDia + "/" + dtMes + "/" + dtAno;
}
