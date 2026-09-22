
		
	function abrir(theURL,winName,features) 
	{ 
	  window.open(theURL,winName,features);
	}

		
	function abre(url)
	{    
				//if(event.srcElement.tagName=='TD')
				//	{
				//		src.children.tags('A')[0].click(); 
				//	}
	location.href = url;

	}	
		
		
	

	function buscaIni(){
	d = document.buscaForm
	d.busca.focus()
	d.busca.select();
	}
	function mudacorBusca(CAMPO){
	CAMPO.style.backgroundColor = '#F5ECBA'
	}
	function mudacorBusca2(CAMPO){
	CAMPO.style.backgroundColor = '#FFFFFF'
	}

 function Verificar()
 {
  var tecla=window.event.keyCode;
  if (tecla==17) {
         buscaIni();
	}
 }
 



function mOvr(src, cor)
    {
    if (!src.contains(event.fromElement))
            {
                src.style.cursor = 'hand';
                src.bgColor = cor;
            }
    }
		
function mOut(src, cor)
{
    if (!src.contains(event.toElement))
        {
            src.style.cursor = 'default';
            src.bgColor = cor;

        }
}