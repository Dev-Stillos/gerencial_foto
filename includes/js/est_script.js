$est = $.noConflict();

$est(document).ready(function(){

    loadPage();
    $est(window).on('hashchange', function() {
      loadPage();
    });
    $est.ajaxSetup({
    cache: false
    });
    $.ajaxSetup({
    cache: false
    });
});

function loadPage(){
   if(location.hash.indexOf('listarItens') > 0){        
        listarItens();
        
    }else if(location.hash.indexOf('cadastrarItens') > 0){        
        cadastrarItens();
        
    }else if(location.hash.indexOf('listarCategorias') > 0){        
        listarCategorias();
            
    }else if(location.hash.indexOf('cadastrarCategorias') > 0){        
        cadastrarCategorias();
            
    }else{   
    
       loadIndex();        
    }
}




// Carregar as paginas
function loadIndex(){
   listarMovimentacoes();
}


//#################################################
// Inicio ---------------------------Item - Relação
//#################################################

function cadastrarRelItens(){
    $.get('includes/asp/est_form_rel_item.asp?data='+Date(),function(data){
        $est('#content').html(data);
       
        $est('#salvar_rel_item').click(function(){
            cadastrarRelItem(1);
        });
        $est('#cancelar_rel_item').click(function(){
            listarRelItens();
        });

    });
}


function cadastrarRelItem(acao){
    if(validaRelItens() || acao == 3){
        $.ajax({
            method: 'post',
            data: 'form=cadastro_rel_item&acao='+acao+'&'+$est('form[name=cadastro_rel_item]').serialize(),
            url: 'includes/asp/est_sql.asp',
            success: function(data){
                if(data.indexOf('Erro')>=0){
                    alerta(data);
                }else{
                    alerta(data);
                    listarRelItens();
                }
            }       
        });
    }
}

function listarRelItens(){
    
    $.get('includes/asp/est_listas.asp?tipo=rel_item&data='+Date(),function(data){
        $est('#content').html(data);

    });
}

function excluirRelItem(id){
    if(confirm('Deseja realmente excluir este registro ?'))
    $.ajax({
        method: 'post',
        data: 'form=cadastro_rel_item&acao=3&cod_rel_item='+id,
        url: 'includes/asp/est_sql.asp',
        success: function(data){
            if(data.indexOf('Erro')>=1){
                alerta(data);
            }else{
                alerta(data);
                listarRelItens();
            }
        }       
    });

}


//#################################################
// Fim ------------------------------Item - Relação
//#################################################



//#################################################
// Inicio ---------------------Item - Relação Venda
//#################################################

function cadastrarRelItensVenda(){
    $.get('includes/asp/est_form_rel_item_venda.asp?data='+Date(),function(data){
        $est('#content').html(data);
       
        $est('#salvar_rel_item').click(function(){
            cadastrarRelItemVenda(1);
        });
        $est('#cancelar_rel_item').click(function(){
            listarRelItensVenda();
        });

    });
}


function cadastrarRelItemVenda(acao){
    if(validaRelItensVenda() || acao == 3){
        $.ajax({
            method: 'post',
            data: 'form=cadastro_rel_item_venda&acao='+acao+'&'+$est('form[name=cadastro_rel_item_venda]').serialize(),
            url: 'includes/asp/est_sql.asp',
            success: function(data){
                if(data.indexOf('Erro')>=0){
                    alerta(data);
                }else{
                    alerta(data);
                    listarRelItensVenda();
                }
            }       
        });
    }
}

function listarRelItensVenda(){
    
    $.get('includes/asp/est_listas.asp?tipo=rel_item_venda&data='+Date(),function(data){
        $est('#content').html(data);

    });
}

function excluirRelItemVenda(id){
    if(confirm('Deseja realmente excluir este registro ?'))
    $.ajax({
        method: 'post',
        data: 'form=cadastro_rel_item_venda&acao=3&cod_rel_item_venda='+id,
        url: 'includes/asp/est_sql.asp',
        success: function(data){
            if(data.indexOf('Erro')>=1){
                alerta(data);
            }else{
                alerta(data);
                listarRelItensVenda();
            }
        }       
    });

}


//#################################################
// Fim ------------------------Item - Relação Venda
//#################################################




//#################################################
// Inicio ----------------------------------Estoque
//#################################################
function listaItensEstoque(){
    
   var select_listar = $est('select[name=select_listar]').val();             
    $.get('includes/asp/est_listas.asp?tipo=estoque&select_listar='+select_listar+'&data='+Date(),function(data){
        $est('#content').html(data);
        $est('select[name=select_listar]').change(function(){
            listaItensEstoque();
            
        });

    });
}

//#################################################
// Fim -------------------------------------Estoque
//#################################################



//#################################################
// Inicio ----------------------------------Pedidos
//#################################################

function listarPedidos(){
    var mes = $est('select[name=mes]').val();
    var ano = $est('select[name=ano]').val();
 
    $.get('includes/asp/est_listas.asp?tipo=pedidos&ano='+ano+'&mes='+mes+'&data='+Date(),function(data){
        $est('#content').html(data);
        $est('select[name=mes] , select[name=ano]').change(function(){
            listarPedidos();
            
        });
    });
}


function cadastrarPedidos(){
    $.get('includes/asp/est_form_pedido.asp?data='+Date(),function(data){
        $est('#content').html(data);
        //$est('select[name=status] option[value=CD]').prop('disabled',true);
        $est('select[name=status] option[value=CN]').prop('disabled',true);
        $est(".datepicker").datepicker({dateFormat: 'dd/mm/yy'});
        $est('.data_pedido').hide();
        $est('#salvar_pedido').click(function(){
            cadastrarPedido(1);
        });
        $est('#cancelar_pedido').click(function(){
            listarPedidos();
        });

    });
}


function addSubItemPedido(){
    if(validateSubItemPedido()){

        cod     = $est('select[name=sub_itens]').val();
        nome    = $est('select[name=sub_itens] option[value='+cod+']').text();
        qtd     = $est('input[name=quantidade_sub_item]').val();
        val     = $est('input[name=valor]').val();
        var item = '<tr id="subItem_'+cod+'" class="item_selecionado"><td>'+nome+'</td><td class="td_quant"> '+qtd+' </td><td> R$ '+val+'</td> <td class="tb_remover"><b class="hide cod">'+cod+'</b><b class="hide qtd">'+qtd+'</b><b class="hide val">'+val+'</b><img onclick="excluirSubItem('+cod+')" class="item_selecionado" src="images/icons/icon_excluir.png"/></td></tr>';

            
        $est('#lista_sub_itens table').append(item);
        loadListSubItemsPedidos();
    }
}

function loadListSubItemsPedidos(){
    var lista = "";
    $est('input[name=list_sub_item]').val("");
    $est('#lista_sub_itens tr').each(function(){
        cod = $est(this).find('b.cod').html();
        qtd = $est(this).find('b.qtd').html();
        val = $est(this).find('b.val').html();
        var valor = ';[['+cod+':'+qtd+':'+val+']]';  
        lista = lista + valor;
    });
    $est('input[name=list_sub_item]').val(lista);
}



function excluirSubItemPedido(cod){
    $est('#subItem_'+cod).remove();
    loadListSubItemsPedidos();
}



function cadastrarPedido(acao){
    if(validaPedido() || acao == 3){
        $.ajax({
            method: 'post',
            data: 'form=cadastro_pedido&acao='+acao+'&'+$est('form[name=cadastro_pedido]').serialize(),
            url: 'includes/asp/est_sql.asp',
            success: function(data){
                if(data.indexOf('Erro')>=0){
                    alerta(data);
                }else{
                    alerta(data);
                    listarPedidos();
                }
            }       
        });
    }
}

function visualizarPedido(id){

    $.get('includes/asp/est_form_pedido.asp?data='+Date(),function(data){
        $est('#content').html(data);
        $est('.titulo_pagina').html('Visualizar item');
        $.post('includes/asp/est_sql.asp',{acao: 4, cod_pedido: id,form:'cadastro_pedido'}, function(data){
            data = $.parseJSON(data);
            $est('input[name=cod_pedido]').val(data.cod_pedido).prop('readonly',true);
            $est('input[name=data_pedido]').val(data.data_pedido).prop('readonly',true);
            $est('input[name=data_previsao]').val(data.data_previsao).prop('readonly',true);
            $est('input[name=data_entrega]').val(data.data_entrega).prop('readonly',true);
            $est('textarea[name=descricao]').val(data.descricao).prop('readonly',true);
            
            $est('select[name=status] option[value='+data.status+']').prop('selected',true).parent().prop('disabled',true);
            $est('select[name=fornecedor] option[value='+data.fornecedor+']').prop('selected',true).parent().prop('disabled',true);

                var sub_itens = data.lista_sub_item.split("]][[");
                
                var lista_sub_itens_form = ""
                for(var i = 0;i < sub_itens.length;i++){
                    itens = sub_itens[i].replace("[[","").replace("]]","");   
                    itens = itens.split("::");
                    cod = itens[0];
                    nome = itens[1];
                    qtd = itens[2];
                    val = itens[3];
                    lista_sub_itens_form += '<tr id="subItem_'+cod+'" class="item_selecionado"><td>'+nome+'</td><td class="td_quant"> '+qtd+' </td><td> R$ '+val+'</td> <td class="tb_remover hide"><b class="hide cod">'+cod+'</b><b class="hide qtd">'+qtd+'</b><b class="hide val">'+val+'</b><img onclick="excluirSubItem('+cod+')" class="item_selecionado" src="images/icons/icon_excluir.png"/></td></tr>';

                   
                               
                }
                $est('#lista_sub_itens table').append(lista_sub_itens_form);
                $est('select[name=sub_itens]').hide();
                $est('input[name=valor]').hide();
                $est('span.moeda').hide();                 
                $est('img.item_selecionado').hide();
                $est('input[name=quantidade_sub_item]').hide();
                $est('#add_sub_item').hide();
                $est('#form_sub_item').show();
                loadListSubItemsPedidos();

        });
        $est('#salvar_pedido').remove();
        $est('#cancelar_pedido').html('Voltar').click(function(){
            
            listarPedidos();
            
         });
         
    });

}

function alterarPedido(id,status){
    if(status != "Em Andamento"){
        alerta('Este item não pode ser alterado.<br>Alterações só são permitidas quando o status do pedido é "Em Andamento"');
    
    }else{
        $.get('includes/asp/est_form_pedido.asp?data='+Date(),function(data){
            $est('#content').html(data);
            $est('.titulo_pagina').html('Alterar item');
                $.post('includes/asp/est_sql.asp',{acao: 4, cod_pedido: id,form:'cadastro_pedido'}, function(data){

                    data = $.parseJSON(data);
                    if(data.status == 'CN' || data.status == 'CD')readonly = true; else readonly = false;
                    $est('input[name=cod_pedido]').val(data.cod_pedido).prop('readonly',readonly);
                    $est('input[name=data_pedido]').val(data.data_pedido).prop('readonly',true);
                    $est('input[name=data_previsao]').val(data.data_previsao).prop('readonly',readonly);
                    $est('input[name=data_entrega]').val(data.data_entrega).prop('readonly',readonly);
                    $est('textarea[name=descricao]').val(data.descricao).prop('readonly',readonly);
                    
                    $est('select[name=status] option[value='+data.status+']').prop('selected',true).parent().prop('disabled',readonly);
                    $est('select[name=fornecedor] option[value='+data.fornecedor+']').prop('selected',true).parent().prop('disabled',readonly);
                    $est(".datepicker").datepicker({dateFormat: 'dd/mm/yy'});
                        var sub_itens = data.lista_sub_item.split("]][[");
                        
                        var lista_sub_itens_form = ""
                        for(var i = 0;i < sub_itens.length;i++){
                            itens = sub_itens[i].replace("[[","").replace("]]","");   
                            itens = itens.split("::");
                            cod = itens[0];
                            nome = itens[1];
                            qtd = itens[2];
                            val = itens[3];
                            lista_sub_itens_form += '<tr id="subItem_'+cod+'" class="item_selecionado"><td>'+nome+'</td><td class="td_quant"> '+qtd+' </td><td> R$ '+val+'</td> <td class="tb_remover"><b class="hide cod">'+cod+'</b><b class="hide qtd">'+qtd+'</b><b class="hide val">'+val+'</b><img onclick="excluirSubItem('+cod+')" class="item_selecionado" src="images/icons/icon_excluir.png"/></td></tr>';                            

                        }
                        $est('#lista_sub_itens table').append(lista_sub_itens_form);
                        if(readonly == true)$est('select[name=sub_itens]').hide();
                        if(readonly == true)$est('input[name=valor]').hide();
                        if(readonly == true)$est('span.moeda').hide();                 
                        if(readonly == true)$est('img.item_selecionado').hide();
                        if(readonly == true)$est('input[name=quantidade_sub_item]').hide();
                        if(readonly == true)$est('#add_sub_item').hide();
                        if(readonly == true)$est('#form_sub_item').show();
                        loadListSubItemsPedidos();
                        if(readonly == true)$est('#salvar_pedido').remove();
                        $est('#cancelar_pedido').html('Voltar').click(function(){
                
                            listarPedidos();
                
                        });
                        $est('#salvar_pedido').click(function(){
            
                            cadastrarPedido(2);
                            
                        });

                });

        
        });
    }
}

//#################################################
// Fim -------------------------------------Pedidos
//#################################################







//#################################################
// Inicio ----------------------------Movimentações
//#################################################
function listarMovimentacoes(){
    
    var mes = $est('select[name=mes]').val();
    var ano = $est('select[name=ano]').val();
    $.get('includes/asp/est_listas.asp?tipo=movimentacoes&ano='+ano+'&mes='+mes+'&data='+Date(),function(data){
        $est('#content').html(data);
        $est('select[name=mes] , select[name=ano]').change(function(){
            listarMovimentacoes();
            
        });
    });
}

var value_anterior_local2 = '';
var value_anterior_local = '';
var value_anterior_tipo = '';
function lancarMovimentacao(){
    $.get('includes/asp/est_form_movimentacao.asp?data='+Date(),function(data){
        $est('#content').html(data);
        $est('#salvar_lancamento_movimentacao').click(function(){
            cadastrarLancamento(1);        
        });
        $est('#cancelar_lancamento_movimentacao').click(function(){
            listarMovimentacoes();
        });
        
        value_anterior_local = $est('select[name=local]').val();
        value_anterior_local2 = $est('select[name=local]').val();
        $est('select[name=local] , select[name=local_origem]').change(function(){
        
            var value = $est('select[name=local]').val();  
            var value2 = $est('select[name=local_origem]').val();             
            
            if($est('input[name=list_sub_item]').val().length < 3){
            
                $est('input[name=list_sub_item]').val("");
                $est('#lista_sub_itens tr').remove();    
                value_anterior_local = $est('select[name=local]').val(); 
                value_anterior_local2 = $est('select[name=local_origem]').val();                
            }
            else if((value_anterior_local2 != value2 || value_anterior_local != value) &&  confirm('Alterando o local, a lista de itens atual será perdida. Deseja continuar?')){
                $est('input[name=list_sub_item]').val("");
                $est('#lista_sub_itens tr').remove();    
                value_anterior_local = $est('select[name=local]').val(); 
                value_anterior_local2 = $est('select[name=local_origem]').val();   
                
            }else{
                $est('select[name=local] option[value='+value_anterior_local+']').prop('selected',true);
                $est('select[name=local_origem] option[value='+value_anterior_local2+']').prop('selected',true);                
            }
            
            
        });
        
        $est('select[name=local] , select[name=local_origem]').click(function(){
           value_anterior_local = $est('select[name=local]').val(); 
           value_anterior_local2 = $est('select[name=local_origem]').val();            
        });

        value_anterior_tipo = $est('select[name=tipo]').val();
        $est('select[name=tipo]').change(function(){
        
            var value = $est('select[name=tipo]').val();   
             if ($est('select[name=tipo]').val() == "IN")
                {
                 
                    $est('select[name=local]').val('0');
                }

            if($est('input[name=list_sub_item]').val().length < 3){
            
                $est('input[name=list_sub_item]').val("");
                $est('#lista_sub_itens tr').remove();    
                value_anterior_tipo = $est('select[name=tipo]').val(); 
            }
            else if(value_anterior_tipo != value &&  confirm('Alterando o tipo de lançamento, a lista de itens atual será perdida. Deseja continuar?')){
                $est('input[name=list_sub_item]').val("");
                $est('#lista_sub_itens tr').remove();    
                value_anterior_tipo = $est('select[name=tipo]').val(); 
                
            }else{
                $est('select[name=tipo] option[value='+value_anterior_tipo+']').prop('selected',true);
            }
            
            
            if(value == 'TRA' || value_anterior_tipo == 'TRA'){
                $('.local_transferencia').show();
                $('.local_origem').html("Local de Origem");
                $('.local_destino').html("Local de Destino");                
            }else{  
                $('.local_transferencia').hide();
                $('.local_destino').html("Local de Armazenamento"); 
            }
            
        });
        
        $est('select[name=tipo]').click(function(){
           value_anterior_tipo = $est('select[name=tipo]').val(); 
        });   
        
        $est('select[name=local]').change(function(){
           var value = $(this).val();
           if(value == -1){
                if ($est('select[name=tipo]').val() == "IN")
                {
                    alert("Atenção a operacao de entrada nao pode ser feita diretamente para um vendedor.\n a entrada deve ser fita ao estoque e depois trasnferido.");
                    $est('select[name=local]').val('1');
                }
                else{
                    $('.select_vendedor').removeClass('hide');
                    $('.select_vendedor').show();
                }
           }else{
               $('.select_vendedor').hide();

           }
        }); 
    });
}

function cadastrarLancamento(acao){
    if(validaLancamento() || acao == 3){

        $.ajax({
            method: 'post',
            data: 'form=lancamento_movimentacao&acao='+acao+'&'+$est('form[name=lancamento_movimentacao]').serialize(),
            url: 'includes/asp/est_sql.asp',
            success: function(data){
                if(data.indexOf('Erro')>=0){
                    alerta(data);
                }else{
                    alerta(data);
                    listarMovimentacoes();
                }
            }       
        });
    }
}

function addSubItemLancamento(){
   validateSubItemLancamento();

}


//#################################################
// Fim -------------------------------Movimentações
//#################################################







//#################################################
// Inicio ------------------------------CRUD Locais
//#################################################

function listarLocais(){

    var select_listar = $est('select[name=select_listar]').val();             
    $.get('includes/asp/est_listas.asp?tipo=locais&select_listar='+select_listar+'&data='+Date(),function(data){
       $est('#content').html(data);
       $est('select[name=select_listar]').change(function(){
            listarLocais();
            
        });
    });
}

function cadastrarLocais(){
    $.get('includes/asp/est_form_local.asp?data='+Date(),function(data){
        $est('#content').html(data);
        $est('#salvar_local').click(function(){
            
            cadastrarLocal(1);
            
         });
         $est('#cancelar_local').click(function(){
            
            listarLocais();
            
         });
         
    });


}

function visualizarLocal(id){

    $.get('includes/asp/est_form_local.asp?data='+Date(),function(data){
        $est('#content').html(data);
        $est('.titulo_pagina').html('Visualizar local');
        $.post('includes/asp/est_sql.asp',{acao: 4, cod_local_armazenamento: id,form:'cadastro_local'}, function(data){
            data = $.parseJSON(data);
            $est('input[name=nome]').val(data.nome).prop('readonly',true);
            $est('textarea[name=descricao]').val(data.descricao).prop('readonly',true);
            $est('input[name=cod_local_armazenamento]').val(data.cod_local_armazenamento).prop('readonly',true);

        });
        $est('#salvar_local').remove();
        $est('#cancelar_local').html('Voltar').click(function(){
            
            listarLocais();
            
         });
         
    });

}


function alterarLocal(id){

    $.get('includes/asp/est_form_local.asp?data='+Date(),function(data){
        $est('#content').html(data);
        $est('.titulo_pagina').html('Alterar local');
        $.post('includes/asp/est_sql.asp',{acao: 4, cod_local_armazenamento: id,form:'cadastro_local'}, function(data){
            data = $.parseJSON(data);
            $est('input[name=nome]').val(data.nome);
            $est('textarea[name=descricao]').val(data.descricao);
            $est('input[name=cod_local_armazenamento]').val(data.cod_local_armazenamento);
            
        });        
        $est('#salvar_local').click(function(){
            
            cadastrarLocal(2);
            
         });
         $est('#cancelar_local').click(function(){
            
            listarLocais();
            
         });
         
    });

}


function cadastrarLocal(acao){
    if(validaCadastroLocal() || acao == 3){

        $.ajax({
            method: 'post',
            data: 'form=cadastro_local&acao='+acao+'&'+$est('form[name=cadastro_local]').serialize(),
            url: 'includes/asp/est_sql.asp',
            success: function(data){
                if(data.indexOf('Erro')>=0){
                    alerta(data);
                }else{
                    alerta(data);
                    listarLocais();
                }
            }       
        });
    }
}

function excluirLocal(id){
    if(confirm('Deseja realmente excluir este registro ?'))
    $.ajax({
        method: 'post',
        data: 'form=cadastro_local&acao=3&cod_local_armazenamento='+id,
        url: 'includes/asp/est_sql.asp',
        success: function(data){
            if(data.indexOf('Erro')>=0){
                alerta(data);
            }else{
                alerta(data);
                listarLocais();
            }
        }       
    });
 
}
//#################################################
// Final -------------------------------CRUD Locais
//#################################################





//#################################################
// Inicio -------------------------------CRUD Itens
//#################################################

function addSubItem(){
    if(validateSubItem()){
        cod     = $est('select[name=sub_itens]').val();
        nome    = $est('select[name=sub_itens] option[value='+cod+']').text();
        qtd     = $est('input[name=quantidade_sub_item]').val();
        var item = '<span id="subItem_'+cod+'" class="item_selecionado">'+nome+', qtd: '+qtd+' <b class="hide cod">'+cod+'</b><b class="hide qtd">'+qtd+'</b><img onclick="excluirSubItem('+cod+')" class="item_selecionado" src="images/icons/icon_excluir.png"/></span>';
        $est('#lista_sub_itens').append(item);
        loadListSubItems();
    }
}

function loadListSubItems(){
    var lista = "";
    $est('input[name=list_sub_item]').val("");
    $est('#lista_sub_itens tr').each(function(){
        cod = $est(this).find('b.cod').html();
        qtd = $est(this).find('b.qtd').html();
        valor = ';[['+cod+':'+qtd+']]';  
        lista = lista + valor;
    });
    $est('input[name=list_sub_item]').val(lista);
}

function excluirSubItem(cod){
    $est('#subItem_'+cod).remove();
    loadListSubItems();
}


function listarItens(){
    
    $.get('includes/asp/est_listas.asp?tipo=itens&data='+Date(),function(data){
        $est('#content').html(data);
    });
}

function cadastrarItens(){
    $.get('includes/asp/est_form_item.asp?data='+Date(),function(data){
        $est('#content').html(data);
         $est('select[name=tipo]').change(function(){
            if($est('select[name=tipo]').val() == 'C'){
                $est('#form_sub_item').show();
            }else{
                $est('#form_sub_item').hide();
            }
        });
        $est('#salvar_item').click(function(){
            
            cadastrarItem(1);
            
         });
         $est('#cancelar_item').click(function(){
            
            listarItens();
            
         });
    });
}

function visualizarItem(id){

    $.get('includes/asp/est_form_item.asp?data='+Date(),function(data){
        $est('#content').html(data);
        $est('.titulo_pagina').html('Visualizar item');
        $.post('includes/asp/est_sql.asp',{acao: 4, cod_item: id,form:'cadastro_item'}, function(data){
            data = $.parseJSON(data);
            $est('input[name=cod_item]').val(data.cod_item).prop('readonly',true);
            $est('input[name=nome]').val(data.nome_item).prop('readonly',true);
            $est('textarea[name=descricao]').val(data.descricao).prop('readonly',true);
            $est('select[name=tipo] option[value='+data.tipo+']').prop('selected',true).parent().prop('disabled',true);
            $est('select[name=cod_categoria] option[value='+data.cod_categoria+']').prop('selected',true).parent().prop('disabled',true);
            if(data.tipo == 'C'){
                var sub_itens = data.lista_sub_item.split("]][[");
                
                var lista_sub_itens_form = ""
                for(var i = 0;i < sub_itens.length;i++){
                    itens = sub_itens[i].replace("[[","").replace("]]","");   
                    itens = itens.split("::");
                    cod = itens[0];
                    nome = itens[1];
                    qtd = itens[2];
                    lista_sub_itens_form += '<span id="subItem_'+cod+'" class="item_selecionado">'+nome+', qtd: '+qtd+' <b class="hide cod">'+cod+'</b><b class="hide qtd">'+qtd+'</b><img onclick="excluirSubItem('+cod+')" class="item_selecionado" src="images/icons/icon_excluir.png"/></span>';
                
                }
                $est('#lista_sub_itens').append(lista_sub_itens_form);
                $est('select[name=sub_itens]').hide();
                $est('img.item_selecionado').hide();
                $est('input[name=quantidade_sub_item]').hide();
                $est('#add_sub_item').hide();
                $est('#form_sub_item').show();
                loadListSubItems();
            }else{
                $est('#form_sub_item').hide();
               
            }
        });
        $est('#salvar_item').remove();
        $est('#cancelar_item').html('Voltar').click(function(){
            
            listarItens();
            
         });
         
    });

}


function alterarItem(id){

    $.get('includes/asp/est_form_item.asp?data='+Date(),function(data){
        $est('#content').html(data);
        $est('.titulo_pagina').html('Alterar item');
        $.post('includes/asp/est_sql.asp',{acao: 4, cod_item: id,form:'cadastro_item'}, function(data){
            data = $.parseJSON(data);
            $est('input[name=cod_item]').val(data.cod_item);
            $est('input[name=nome]').val(data.nome_item);
            $est('textarea[name=descricao]').val(data.descricao);
            $est('select[name=tipo] option[value='+data.tipo+']').prop('selected',true);
            $est('select[name=cod_categoria] option[value='+data.cod_categoria+']').prop('selected',true);
            if(data.tipo == 'C'){
                var sub_itens = data.lista_sub_item.split("]][[");
                
                var lista_sub_itens_form = ""
                for(var i = 0;i < sub_itens.length;i++){
                    itens = sub_itens[i].replace("[[","").replace("]]","");   
                    itens = itens.split("::");
                    cod = itens[0];
                    nome = itens[1];
                    qtd = itens[2];
                    lista_sub_itens_form += '<span id="subItem_'+cod+'" class="item_selecionado">'+nome+', qtd: '+qtd+' <b class="hide cod">'+cod+'</b><b class="hide qtd">'+qtd+'</b><img onclick="excluirSubItem('+cod+')" class="item_selecionado" src="images/icons/icon_excluir.png"/></span>';
                
                }
                $est('#lista_sub_itens').append(lista_sub_itens_form);
                $est('#form_sub_item').show();
                loadListSubItems();
            }else{
                $est('#form_sub_item').hide();
               
            }
            
        });
         $est('select[name=tipo]').change(function(){
            if($est('select[name=tipo]').val() == 'C'){
                $est('#form_sub_item').show();
            }else{
                $est('#form_sub_item').hide();
            }
        });
        
        $est('#salvar_item').click(function(){
            
            cadastrarItem(2);
            
         });
         $est('#cancelar_item').click(function(){
            
            listarItens();
            
         });
         
    });

}


function cadastrarItem(acao){
    if(validaCadastroItem() || acao == 3){

        $.ajax({
            method: 'post',
            data: 'form=cadastro_item&acao='+acao+'&'+$est('form[name=cadastro_item]').serialize(),
            url: 'includes/asp/est_sql.asp',
            success: function(data){
                if(data.indexOf('Erro')>=0){
                    alerta(data);
                }else{
                    alerta(data);
                    listarItens();
                }
            }       
        });
    }
}

function excluirItem(id){
    if(confirm('Deseja realmente excluir este registro ?'))
    $.ajax({
        method: 'post',
        data: 'form=cadastro_item&acao=3&cod_item='+id,
        url: 'includes/asp/est_sql.asp',
        success: function(data){
            if(data.indexOf('Erro')>=1){
                alerta(data);
            }else{
                alerta(data);
                listarItens();
            }
        }       
    });
 
}
//#################################################
// Final --------------------------------CRUD Itens
//#################################################



//#################################################
// Inicio --------------------------CRUD Categorias
//#################################################


function listarCategorias(){
    $.get('includes/asp/est_listas.asp?tipo=categorias&data='+Date(),function(data){
        $est('#content').html(data);
    });
}

function cadastrarCategorias(){
    $.get('includes/asp/est_form_categoria.asp?data='+Date(),function(data){
        $est('#content').html(data);
      
        $est('#salvar_categoria').click(function(){
            
            cadastrarCategoria(1);
            
         });
         $est('#cancelar_categoria').click(function(){
            
            listarCategorias();
            
         });
    });
}

function visualizarCategoria(id){

    $.get('includes/asp/est_form_categoria.asp?data='+Date(),function(data){
        $est('#content').html(data);
        $est('.titulo_pagina').html('Visualizar categoria');
        $.post('includes/asp/est_sql.asp',{acao: 4, cod_categoria: id,form:'cadastro_categoria'}, function(data){
            console.log(data);
            data = $.parseJSON(data);
            
            $est('input[name=cod_categoria]').val(data.cod_categoria).prop('readonly',true);
            $est('input[name=nome]').val(data.nome).prop('readonly',true);
            $est('textarea[name=descricao]').val(data.descricao).prop('readonly',true);            
        });
        $est('#salvar_categoria').remove();
        $est('#cancelar_categoria').html('Voltar').click(function(){
            
            listarCategorias();
            
         });
         
    });

}


function alterarCategoria(id){

    $.get('includes/asp/est_form_categoria.asp?data='+Date(),function(data){
        $est('#content').html(data);
        $est('.titulo_pagina').html('Alterar categoria');
        $.post('includes/asp/est_sql.asp',{acao: 4, cod_categoria: id,form:'cadastro_categoria'}, function(data){
            data = $.parseJSON(data);
            $est('input[name=cod_categoria]').val(data.cod_categoria);
            $est('input[name=nome]').val(data.nome);
            $est('textarea[name=descricao]').val(data.descricao);            

        });

        $est('#salvar_categoria').click(function(){
            
            cadastrarCategoria(2);
            
         });
         $est('#cancelar_categoria').click(function(){
            
            listarCategorias();
            
         });
         
    });

}


function cadastrarCategoria(acao){
    if(validaCadastroCategoria() || acao == 3){

        $.ajax({
            method: 'post',
            data: 'form=cadastro_categoria&acao='+acao+'&'+$est('form[name=cadastro_categoria]').serialize(),
            url: 'includes/asp/est_sql.asp',
            success: function(data){
                if(data.indexOf('Erro')>=0){
                    alerta(data);
                }else{
                    alerta(data);
                    listarCategorias();
                }
            }       
        });
    }
}

function excluirCategoria(id){
    if(confirm('Deseja realmente excluir este registro ?'))
    $.ajax({
        method: 'post',
        data: 'form=cadastro_categoria&acao=3&cod_categoria='+id,
        url: 'includes/asp/est_sql.asp',
        success: function(data){
            if(data.indexOf('Erro')>=1){
                alerta(data);
            }else{
                alerta(data);
                listarCategorias();
            }
        }       
    });
 
}
//#################################################
// Final ---------------------------CRUD Categorias
//#################################################



//#################################################
// Inicio ------------------------CRUD Fornecedores
//#################################################


function listarFornecedores(){
    $.get('includes/asp/est_listas.asp?tipo=fornecedores&data='+Date(),function(data){
        $est('#content').html(data);
    });
}

function cadastrarFornecedores(){
    $.get('includes/asp/est_form_fornecedor.asp?data='+Date(),function(data){
        $est('#content').html(data);
      
        $est('#salvar_fornecedor').click(function(){
            
            cadastrarFornecedor(1);
            
         });
         $est('#cancelar_fornecedor').click(function(){
            
            listarFornecedores();
            
         });
    });
}

function visualizarFornecedor(id){

    $.get('includes/asp/est_form_fornecedor.asp?data='+Date(),function(data){
        $est('#content').html(data);
        $est('.titulo_pagina').html('Visualizar fornecedor');
        $.post('includes/asp/est_sql.asp',{acao: 4, cod_fornecedor: id,form:'cadastro_fornecedor'}, function(data){
            data = $.parseJSON(data);
            $est('input[name=cod_fornecedor]').val(data.cod_fornecedor).prop('readonly',true);
            $est('input[name=nome]').val(data.nome).prop('readonly',true);
            $est('textarea[name=descricao]').val(data.descricao).prop('readonly',true);            
        });
        $est('#salvar_fornecedor').remove();
        $est('#cancelar_fornecedor').html('Voltar').click(function(){
            
            listarFornecedores();
            
         });
         
    });

}


function alterarFornecedor(id){

    $.get('includes/asp/est_form_fornecedor.asp?data='+Date(),function(data){
        $est('#content').html(data);
        $est('.titulo_pagina').html('Alterar fornecedor');
        $.post('includes/asp/est_sql.asp',{acao: 4, cod_fornecedor: id,form:'cadastro_fornecedor'}, function(data){
            data = $.parseJSON(data);
            $est('input[name=cod_fornecedor]').val(data.cod_fornecedor);
            $est('input[name=nome]').val(data.nome);
            $est('textarea[name=descricao]').val(data.descricao);            

        });

        $est('#salvar_fornecedor').click(function(){
            
            cadastrarFornecedor(2);
            
         });
         $est('#cancelar_fornecedor').click(function(){
            
            listarFornecedores();
            
         });
         
    });

}


function cadastrarFornecedor(acao){

    if(validaCadastroFornecedor() || acao == 3){

        $.ajax({
            method: 'post',
            data: 'form=cadastro_fornecedor&acao='+acao+'&'+$est('form[name=cadastro_fornecedor]').serialize(),
            url: 'includes/asp/est_sql.asp',
            success: function(data){
                if(data.indexOf('Erro')>=0){
                    alerta(data);
                }else{
                    alerta(data);
                    listarFornecedores();
                }
            }       
        });
    }
}

function excluirFornecedor(id){
    if(confirm('Deseja realmente excluir este registro ?'))
    $.ajax({
        method: 'post',
        data: 'form=cadastro_fornecedor&acao=3&cod_fornecedor='+id,
        url: 'includes/asp/est_sql.asp',
        success: function(data){
            if(data.indexOf('Erro')>=0){
                alerta(data);
            }else{
                alerta(data);
                listarFornecedores();
            }
        }       
    });
 
}
//#################################################
// Final -------------------------CRUD Fornecedores
//#################################################




//#################################################
// Inicio ---------------Validações dos formulários
//#################################################

function validaLancamento(){
var validate = true;
var erroMsg = '';

    if(!validaSelect($est('select[name=tipo]'))){validate = false;erroMsg += 'Selecione um tipo de lançamento.<br>';}
    if(!validaSelect($est('select[name=local]'))){validate = false;erroMsg += 'Selecione um local de armazenamento.<br>';}
    if($est('select[name=local]').val() == -1)if(!validaSelect($est('select[name=vendedor]'))){validate = false;erroMsg += 'Selecione um vendedor.<br>';}
    if($est('input[name=list_sub_item]').val().length <2 ){validate = false;erroMsg += 'Selecione um item para movimentação.<br>';}    

    if(erroMsg.length > 0)alerta(erroMsg);
    
return validate;
}


function validaPedido(){
var validate = true;
var erroMsg = '';
    if($est('select[name=status]').val() == 'CD')entrega = true; else entrega = false;
    if(!validaSelect($est('select[name=status]'))){validate = false;erroMsg += 'Selecione um status.<br>';}
    if(!validaSelect($est('select[name=fornecedor]'))){validate = false;erroMsg += 'Selecione um fornecedor.<br>';}
    if($est('input[name=list_sub_item]').val().length <2 ){validate = false;erroMsg += 'Selecione no mínimo um item para o pedido.<br>';}  
    if(!validaDate($est('input[name=data_previsao]'),true)){validate = false;erroMsg += 'Preencha o campo data de previsão corretamente.<br>';}
    if(!validaDate($est('input[name=data_entrega]'),entrega)){validate = false;erroMsg += 'Preencha o campo data de entrega corretamente.<br>';}
        

    if(erroMsg.length > 0)alerta(erroMsg);
    
return validate;
}


function validateSubItemLancamento(){
  
var validate = true;
var erroMsg = '';

    

    if(!validaSelect($est('select[name=sub_itens]'))){
    
        validate = false;
        erroMsg += 'Selecione um item.<br>';
    }
    if(!validaNumber($est('input[name=quantidade_sub_item]'),1,true)){
        validate = false;
        erroMsg += 'Selecione a quantia de itens.(qtd > 0) <br>';
    
    }  

    if($est('input[name=list_sub_item]').val().indexOf('[['+$est('select[name=sub_itens]').val()+':')>=0){
        validate = false;
        erroMsg += 'Este item já foi adicionado à lista. <br>';
    
    }      
    
    if(validate == true){
        return validateQuantiaItensLancamento();
    }else{
        if(erroMsg.length > 0)alerta(erroMsg); 
        return validate;        
    }
    
}


function validateQuantiaItensLancamento(){

 var validate = false;
 var erroMsg = "";
 var lista_itens = '';
 
 if($est('select[name=tipo]').val()== "OUT" || $est('select[name=tipo]').val()== "TRA"){
        var cod_item = $est('select[name=sub_itens]').val();
        var qtd = $est('input[name=quantidade_sub_item]').val();
        var local = "";
        var cod_vendedor = $est('vendedor').val();
        if($est('select[name=tipo]').val()== "OUT")local = $est('select[name=local]').val() ; else local = $est('select[name=local_origem]').val() ;
        $.ajax({
            method: 'post',
            data: 'form=lancamento_movimentacao&acao=5&cod_local='+local+'&cod_item='+cod_item+'&quantia='+qtd+'&vendedor='+cod_vendedor,
            url: 'includes/asp/est_sql.asp',
            success: function(data){
                
                if(data.length<=1){
                
                    cod     = $est('select[name=sub_itens]').val();
                    nome    = $est('select[name=sub_itens] option[value='+cod+']').text();
                    qtd     = $est('input[name=quantidade_sub_item]').val();
                      var item = '<tr id="subItem_'+cod+'" class="item_selecionado"><td>'+nome+'</td><td class="td_quant"> '+qtd+' </td> <td class="tb_remover"><b class="hide cod">'+cod+'</b><b class="hide qtd">'+qtd+'</b><img onclick="excluirSubItem('+cod+')" class="item_selecionado" src="images/icons/icon_excluir.png"/></td></tr>';
                    $est('#lista_sub_itens table').append(item);
                    loadListSubItems();
                
                }else{
                
                    if(erroMsg.length<2)erroMsg += data;
                    alerta(erroMsg);
                
                }
                
            }       
        });
    }else{
    
        cod     = $est('select[name=sub_itens]').val();
        nome    = $est('select[name=sub_itens] option[value='+cod+']').text();
        qtd     = $est('input[name=quantidade_sub_item]').val();
        var item = '<tr id="subItem_'+cod+'" class="item_selecionado"><td>'+nome+'</td><td class="td_quant"> '+qtd+' </td> <td class="tb_remover"><b class="hide cod">'+cod+'</b><b class="hide qtd">'+qtd+'</b><img onclick="excluirSubItem('+cod+')" class="item_selecionado" src="images/icons/icon_excluir.png"/></td></tr>';
        $est('#lista_sub_itens table').append(item);
        loadListSubItems();        
    
    }
}


function validaRelItens(){
var validate = true;
var erroMsg = '';

    if(!validaSelect($est('select[name=cod_item_p]'))){validate = false;erroMsg += 'Selecione um item (Produção).<br>';}
    if(!validaSelect($est('select[name=cod_item_e]'))){validate = false;erroMsg += 'Selecione um item (Estoque). <br>';}
    if(!validaSelect($est('select[name=cod_local]'))){validate = false;erroMsg += 'Selecione um local. <br>';}
    
    if(erroMsg.length > 0)alerta(erroMsg);
    
    
return validate;

}


function validaRelItensVenda(){
var validate = true;
var erroMsg = '';

    if(!validaSelect($est('select[name=cod_item_p]'))){validate = false;erroMsg += 'Selecione um item (Venda).<br>';}
    if(!validaSelect($est('select[name=cod_item_e]'))){validate = false;erroMsg += 'Selecione um item (Estoque). <br>';}
    if(!validaSelect($est('select[name=cod_local]'))){validate = false;erroMsg += 'Selecione um local. <br>';}
    if(erroMsg.length > 0)alerta(erroMsg);
    
    
return validate;

}

function validateSubItem(){

var validate = true;
var erroMsg = '';

    if(!validaSelect($est('select[name=sub_itens]'))){validate = false;erroMsg += 'Selecione um item.<br>';}
    if(!validaNumber($est('input[name=quantidade_sub_item]'),1,true)){validate = false;erroMsg += 'Selecione a quantia de itens.(qtd > 0) <br>';}
    if(erroMsg.length > 0)alerta(erroMsg);
    
return validate;

}


function validateSubItemPedido(){

var validate = true;
var erroMsg = '';

    if(!validaSelect($est('select[name=sub_itens]'))){validate = false;erroMsg += 'Selecione um item.<br>';}
    if(!validaNumber($est('input[name=quantidade_sub_item]'),1,true)){validate = false;erroMsg += 'Selecione a quantia de itens.(qtd > 0) <br>';}
    if(!validaNumber($est('input[name=valor]'),1,true)){validate = false;erroMsg += 'Preencha o campo valor do item corretamente.<br>';}
    if($est('input[name=list_sub_item]').val().indexOf('[['+$est('select[name=sub_itens]').val()+':') >=0){
        validate = false;
        erroMsg += 'Este item já foi adicionado à lista. <br>';
    }    
    if(erroMsg.length > 0)alerta(erroMsg);
    
return validate;

}


function validaCadastroItem(){
var validate = true;
var erroMsg = '';

    if(!validaText($est('input[name=nome]'),3,true)){validate = false;erroMsg += 'Preencha o campo nome corretamente.<br>';}
    if(!validaSelect($est('select[name=cod_categoria]'))){validate = false;erroMsg += 'Selecione uma categoria.<br>';}
    if(!validaSelect($est('select[name=tipo]'))){validate = false;erroMsg += 'Selecione um tipo de item.<br>';}
    if($est('select[name=tipo]').val() == "C"){if($est('input[name=list_sub_item]').val().length <2 ){validate = false;erroMsg += 'Selecione um sub item.<br>';}}    

    if(erroMsg.length > 0)alerta(erroMsg);
    
return validate;
}



function validaCadastroLocal(){
var validate = true;
var erroMsg = '';

    if(!validaText($est('input[name=nome]'),1,true)){validate = false;erroMsg += 'Preencha o campo nome corretamente.<br>';}

    if(erroMsg.length > 0)alerta(erroMsg);
    
return validate;
}


function validaCadastroCategoria(){
var validate = true;
var erroMsg = '';

    if(!validaText($est('input[name=nome]'),1,true)){validate = false;erroMsg += 'Preencha o campo nome corretamente.<br>';}

    if(erroMsg.length > 0)alerta(erroMsg);
    
return validate;
}

function validaCadastroFornecedor(){
var validate = true;
var erroMsg = '';

    if(!validaText($est('input[name=nome]'),1,true)){validate = false;erroMsg += 'Preencha o campo nome corretamente.<br>';}

    if(erroMsg.length > 0)alerta(erroMsg);
    
return validate;
}

//#################################################
// Fim ------------------Validações dos formulários
//#################################################
