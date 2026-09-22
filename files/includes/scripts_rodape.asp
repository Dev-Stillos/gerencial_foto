<script type="text/javascript">

function padraoEntradaFocus(){

	$('input:text,select,textarea').focus(function() {
		var e = $(this).attr("type");
		var a = $(this).hasClass('inputContaCaracteres');
		
		if (e !== 'submit' && !a) {
	
			$(this).css("box-shadow","0px 0px 2px #666");
			$(this).addClass("eShadow");
		 
			if(!$(this).hasClass('formNaoValidado')){
				//$(this).css("border","1px solid #5555fa");
				
				$(this).css("border-left","1px solid #5555fa");
				$(this).css("border-bottom","1px solid #00008b");
				$(this).css("border-top","1px solid #5555fa");
				$(this).css("border-right","1px solid #00008b");
				
				//$(this).css("background-color","ebf4f7");
				$(this).css("background-color","#faf9f9");
			}
		}	 
	});
}

function padraoSaidaFocus(){
	
	$('input:text,select,textarea').blur(function() {
		
		var e = $(this).attr("type");
		var a = $(this).hasClass('inputContaCaracteres');
		
		if (e !== 'submit' && !a) {
		
			$(this).css("box-shadow","none");
			$(this).removeClass("eShadow");
			
			if(!$(this).hasClass('formNaoValidado')){
				$(this).css("border","1px solid #666666");
				$(this).css("background-color","#F4F4F4");
				
			}
		}
	});

	$('.padrao_valor,.padrao_quantidade_decimal').blur(function() {
		var temp = $(this).val();
		 if( temp == '0,00' || temp == '0,000'){
			$(this).val('');
		 }
	});
}

setTimeout("efeitoFormataFocus()",1);

function efeitoFormataFocus(){
		padraoEntradaFocus();
		padraoSaidaFocus();

		//Mascaras novas
		jQuery(function($){

		   $('.padrao_cpf').mask("999.999.999-99",{placeholder:'_'});
		   $('.padrao_cnpj').mask("99.999.999/9999-99",{placeholder:'_'});
		   $('.padrao_data').mask("99/99/9999",{placeholder:'_'});
		   $('.padrao_hora').mask("99:99:99",{placeholder:'_'});
		   $('.padrao_cep').mask("99999-999",{placeholder:'_'});
		   $('.padrao_telefone').mask("(99) 9999-9999",{placeholder:'_'});
		   $('.padrao_celular').mask("(99) 9999-9999?9",{placeholder:'_'});
		   $('.padrao_numero').mask("?9999999999999999",{placeholder:''});
		   $('.ano_conclusao').mask("99/9999",{placeholder:''});

		});
		
		$(function(){
			$('.padrao_valor').priceFormat({
				prefix: '',
				centsSeparator: ',',
				thousandsSeparator: '.'
			}); 
		});

		$(function(){
			$('.padrao_quantidade_decimal').priceFormat({
				prefix: '',
				centsSeparator: ',',
				thousandsSeparator: '.',
				centsLimit: 3
			}); 
		});		
}
</script>