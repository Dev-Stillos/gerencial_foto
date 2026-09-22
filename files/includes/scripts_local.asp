<script type="text/javascript">

function retornaValorBuscaCampo(valor,campo){
	campo.value = valor
	campo.focus();
}

function iniciarBuscaParaCampo(e,valor,tipo,formulario){
	var whichCode = (window.event)?event.keyCode:e.which; 
	if(whichCode == 13){							
		var caminho = "buscaParaCampo_ajax.asp?valor="+valor+"&campo="+tipo+"&formulario="+formulario;
		ajax_loadContent('listagemResultadoBusca',caminho,false,false,false);
	}
}
</script>
<style type="text/css">
<!--
	#ajax_listOfOptions{
		position:absolute;	/* Never change this one */
		width:419px;	/* Width of box */
		height:200px;	/* Height of box */
		overflow:auto;	/* Scrolling features */
		border:1px solid #317082;	/* Dark green border */
		background-color:#FFF;	/* White background color */
		text-align:left;
		font-size:0.9em;
		z-index:100;
	}
	#ajax_listOfOptions div{	/* General rule for both .optionDiv and .optionDivSelected */
		margin:1px;		
		padding:1px;
		cursor:pointer;
		font-size:11px;
		position:relative;
		float:left;
		border-bottom:solid 1px #F0F0F0;
	}
	#ajax_listOfOptions .optionDiv{	/* Div for each item in list */
	}
	
	#ajax_listOfOptions .optionDiv strong{
		position:relative;
		float:left;
		width:36px;
		font-weight:normal;
	}

	#ajax_listOfOptions .optionDiv span{
		position:relative;
		float:left;
		width:299px;
	}

	#ajax_listOfOptions .optionDiv em{
		position:relative;
		float:left;
		width:60px;
		font-style:normal;
	}
	
	#ajax_listOfOptions .optionDivSelected{ /* Selected item in the list */
		background-color:#266493;
		text-decoration:none;
		color:#FFFFFF;
	}
	
	#ajax_listOfOptions .optionDivSelected strong{
		position:relative;
		float:left;
		width:36px;
		font-weight:normal;
	}

	#ajax_listOfOptions .optionDivSelected span{
		position:relative;
		float:left;
		width:299px;
	}

	#ajax_listOfOptions .optionDivSelected em{
		position:relative;
		float:left;
		width:60px;
		font-style:normal;
	}
	#ajax_listOfOptions_iframe{
		background-color:#F00;
		position:absolute;
		z-index:5;
	}
-->
</style>