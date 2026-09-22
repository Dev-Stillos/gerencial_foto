var dynamicContent_ajaxObjects = new Array();
var retornarAjax = null;
//Função que somente processa o ajax e retorna null ou funções, mas não preenche espaços de DIVs
function ajax_processa(caminho,telaCarregar){
	
	if(telaCarregar){
		displayCarregando();
	}
	
	if (window.XMLHttpRequest){
	  xmlhttp=new XMLHttpRequest();
	}else {
	  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
  	}	
	
	xmlhttp.onreadystatechange=function() {
		 if(xmlhttp.readyState==4 && xmlhttp.status==200){
						
			if(telaCarregar){
				displayCarregandoClose();
			}
				
			var valorResultado = new String(xmlhttp.responseText);
				
			if(valorResultado.indexOf('erro#') == 0){
				arrayPrincipal = valorResultado.split('#')
		
				//Executa tarefas, verifica se está habilitado	
				if(arrayPrincipal[1] !== ""){
					eval(arrayPrincipal[1]);
				}
			}else{
			arrayPrincipal = valorResultado.split('#')
		
				//Parametros de sucesso
				if(arrayPrincipal[1] == "true" ){
					eval(arrayPrincipal[2]);
				}		
			}
    	}
  	}
xmlhttp.open("GET",caminho,true);
xmlhttp.send();
}

function ajax_retorna(url)
{	
	var ajaxIndex = dynamicContent_ajaxObjects.length;		
	dynamicContent_ajaxObjects[ajaxIndex] = new sack();
	dynamicContent_ajaxObjects[ajaxIndex].requestFile = url;	// Specifying which file to get	
	dynamicContent_ajaxObjects[ajaxIndex].onCompletion = function(){ 
				
				var valorResultado = new String(dynamicContent_ajaxObjects[ajaxIndex].response);
				self.retornarAjax = valorResultado;
				
				
				if(valorResultado.indexOf('erro#') == 0){
					arrayPrincipal = valorResultado.split('#')
					//Executa tarefas, verifica se está habilitado	
					if(arrayPrincipal[1] !== ""){
							eval(arrayPrincipal[1]);
					}
				}else{
					
					arrayPrincipal = valorResultado.split('#')
					//Parametros de sucesso
					if(arrayPrincipal[1] == "true" ){
						eval(arrayPrincipal[2]);
					}
					
				}
				
				ajax_retornarAjax(valorResultado);
				dynamicContent_ajaxObjects[ajaxIndex] = false; 
			
			};	// Specify function that will be executed after file has been found	
	dynamicContent_ajaxObjects[ajaxIndex].runAJAX();		// Execute AJAX function
	var aux = retornarAjax;
	self.retornarAjax = null;
	if(aux){
		return aux;
	}else{
		return false;
	}
}


//Função que processa e exibe em DIV
function ajax_showContent(divId,ajaxIndex,url)
{
	var valorResultado = new String(dynamicContent_ajaxObjects[ajaxIndex].response);
	var targetObj = document.getElementById(divId);
	
	if(valorResultado.indexOf('erro#') == 0){
		arrayPrincipal = valorResultado.split('#')
				
		//Executa tarefas, verifica se está habilitado	
		if(arrayPrincipal[1] !== ""){
				eval(arrayPrincipal[1]);
		}
		
	}else{
		//Mostra o conteúdo
		//document.write(valorResultado);
		targetObj.innerHTML = valorResultado;	
		
		arrayPrincipal = valorResultado.split('#')
		//Parametros de sucesso
		if(arrayPrincipal[1] == "true" ){
			if(arrayPrincipal[2]){
				eval(arrayPrincipal[2]);
			}
		}		
	}
	
	dynamicContent_ajaxObjects[ajaxIndex] = false; 
}

function ajax_loadContent(divId,url,telaCarregar,fecharTelaCarregar,usarCache)
{			
			
			
			//usarCache variável não está sendo mais utilizada
			
			var ajaxIndex = dynamicContent_ajaxObjects.length;
	
			if(telaCarregar){
				displayCarregando();
			}
		
		dynamicContent_ajaxObjects[ajaxIndex] = new sack();
		dynamicContent_ajaxObjects[ajaxIndex].requestFile = url;	// Specifying which file to get
		dynamicContent_ajaxObjects[ajaxIndex].onCompletion = function(){ 
			
			if(telaCarregar){
				displayCarregandoClose();
			}
		ajax_showContent(divId,ajaxIndex,url); 
		efeitoFormataFocus();
		
	};	// Specify function that will be executed after file has been found
	dynamicContent_ajaxObjects[ajaxIndex].runAJAX();		// Execute AJAX function	
	
}
	