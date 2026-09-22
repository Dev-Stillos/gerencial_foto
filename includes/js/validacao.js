//#######################################################
//Carregar Formulários
//#######################################################
function validaText(who,min_size,required){
    $est(who).css('background-color','');
    if(required == true)
       if($est(who).val().length < min_size)return errorField(who); else return true;
    else
       if($est(who).val().length > 0 && $est(who).val().length < min_size)return errorField(who); else return true;
}

function validaNumber(who,min_size,required){
    $est(who).css('background-color','');
    var val = $est(who).val().replace(',','').replace('.','');
    if(required == true)
       if(val < min_size || isNaN(Number(val)))return errorField(who); else return true;
    else
       if(val.length > 0 && val < min_size  || isNaN(Number(val)))return errorField(who); else return true;
}

function validaEmail(who,required) {
    $est(who).css('background-color',''); 
    var email = $est(who).val();
    var re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    if(required == true)
       if(!re.test(email))return errorField(who); else return true;
    else
       if(!re.test(email) && $est(who).val().length > 0)return errorField(who); else return true;   
} 

function validaDate(who,required){
    $est(who).css('background-color',''); 
    if(required == true){
        var text = $est(who).val();
        var comp = text.split('/');
        var m = parseInt(comp[1], 10);
        var d = parseInt(comp[0], 10);
        var y = parseInt(comp[2], 10);
        var date = new Date(y,m-1,d);
        if (date.getFullYear() == y && date.getMonth() + 1 == m && date.getDate() == d) {
        
            return true;
            
        }else{
        
            return errorField(who); 
           
        }
        
    }else{
        if($est(who).val().length >= 1){
            var text = $est(who).val();
            var comp = text.split('/');
            var m = parseInt(comp[1], 10);
            var d = parseInt(comp[0], 10);
            var y = parseInt(comp[2], 10);
            var date = new Date(y,m-1,d);
            if (date.getFullYear() == y && date.getMonth() + 1 == m && date.getDate() == d) {
            
                return true;
                
            }else{
            
                return errorField(who); 
               
            }
        }else{
            return true;
        }
    
    }
}
function validaCheckBox(who){
    $est(who).css('background-color','');
    if(!$est(who).is(':checked'))return errorField(who); else return true;
}

function validaSelect(who){
    $est(who).css('background-color','');
    if($est(who).val() == 0)return errorField(who); else return true;
}

function errorField(who){
    $est(who).focus().css('background-color','#FFC2C2');
    return false;
}

function alerta(msg){
	 alert(msg);
	 /*
    $est('#alert').remove();
    $est('body').append('<div id="alert"><p>    Preencha o campo corretamente    </p>  <span id="close_alert"> OK </span>');
   
    $est('#alert p').html(msg).parent().show();
    width = $est('#alert').width();
    width_window = $( window ).width();
    h_pos = (width_window/2)-(width/2);
    $est('#alert').css('margin-left',''+h_pos+'px');
    $est('#close_alert').click(function(){$est('#alert').fadeOut(500).remove();});
    */
}
function onLoad(){
    $est('#load').remove();
    $est('body').append('<div id="load"><p>Sua adesão esta sendo processada.<br>Por favor aguarde.</p>');
    $est('#load').show();
}
function afterLoad(){
    $est('#load').fadeOut(500).remove();
}
