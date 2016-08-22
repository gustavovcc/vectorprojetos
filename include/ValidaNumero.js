function Valida(parValor) {	
	var string = parValor;
	var posicao = string.indexOf(".");
	var dados = "";
	if (posicao != -1) {
		tamanho = (parValor.length);
		for(i=0;i<=tamanho;i++) {
			if (i != posicao) {
				dados = dados + string.substring(i,i+1);
			}
		}
	}
	else {
		dados = parValor;
	}
	flag = ValidaNumero(dados);
	return flag;
}	
function ValidaNumero(parNumero) {
	var l_js_Contador; 
	var l_js_Retorno = true;
	for (l_js_Contador = 0; l_js_Contador < parNumero.length; l_js_Contador++) {
		var l_js_Char = parNumero.charAt(l_js_Contador);
		if (!((l_js_Char >= "0") && (l_js_Char <= "9")))
			l_js_Retorno = false;
	}
	return l_js_Retorno;
}

