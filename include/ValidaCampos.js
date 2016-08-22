	function ValidaDados() {
		if (frmDados.txtLocalidade.value == "") {
			alert("Preencha o campo Localidade.");
			frmDados.txtLocalidade.focus();
			return false;
		}
		if (frmDados.txtLocalidade.value != "") {
			if (Valida(frmDados.txtLocalidade.value) == false) {
				alert("O campo Localidade não é válido.");
				frmDados.txtLocalidade.focus();
				return false;
			}
		}
			if (confirm("Deseja salvar os dados?")) {
				frmDados.action = "prog/baixo.asp?Enviar=S";
				frmDados.submit();
		}
		return false; 
	}
