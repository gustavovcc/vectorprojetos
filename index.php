<?php
 #abaixo, criamos uma variavel que ter� como conte�do o endere�o para onde haver� o redirecionamento:  
 $redirect = "http://186.215.108.111:90";
 
 #abaixo, chamamos a fun��o header() com o atributo location: apontando para a variavel $redirect, que por 
 #sua vez aponta para o endere�o de onde ocorrer� o redirecionamento
 header("location:$redirect");
 
?>