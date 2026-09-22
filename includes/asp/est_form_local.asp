<!-- #include file="../../includes/xqrtopqqetyynmkgfdconexao.asp" -->
<!-- #include file="../../includes/asp/valida_usuario.asp" -->
<form name="cadastro_local" class="formulario cadastro_local">
    
    <h1 class="titulo_pagina" > Cadastro de Locais de Armazenamento </h1>
    <table border="1">
        <tr>
            <input type="hidden" name="cod_local_armazenamento" value="0"/>
            <td class="coluna_formulario_label">
                <label> Nome </label>
            </td>        
            <td class="coluna_formulario_fields">
                <input type="text" name="nome"/>
            </td>
            
        </tr>    

              <tr>
             <td class="coluna_formulario_label">
                <label> Descrição </label>
            </td>    
            <td class="coluna_formulario_fields">
                <textarea name="descricao"></textarea><br>
            </td>
        </tr>    
        <tr>
            <td  colspan="2" class="formulario_botoes">

                <span id="salvar_local"> Salvar </span>
                <span id="cancelar_local"> Cancelar </span>                
                
            </td>
        </tr>  
    </table>    
</form>