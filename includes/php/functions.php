<?php
/*Biblioteca de Funcoes padrões*/

@session_start();


/*Criando variaveis variaveis a partir dos dados retornados:
 *
 *Dados dos verbos: SESSION, POST e GET
 *          Ex: url?acao=1&valor=50
 *          
 *Criada as seguintes variaveis:
 *          $acao=1 e $valor=50
 */
foreach ($_POST as $key=>$value)
{
   ${$key}=$value;
}

foreach ($_GET as $key=>$value)
{
   ${$key}=$value;
}

foreach ($_SESSION as $key=>$value)
{
   ${"session_".$key}=$value;
}


/*Funções de auxilio a formatação das informações
 *para exibição e inserção no banco de dados.
 */

function sql_date_format($field,$return_name = "",$format = "103"){
  if($return_name == "")$return_name = $field;
  return " CONVERT(VARCHAR(25),$field,$format) AS $return_name ";
}

function rbl($value){
    return @number_format($value,2,',','.');
}

function rbl_to_insert($value){
    return @number_format($value,2,'.','');    
}

function date_compare($start_date,$end_date){
   $return = 0;
   $start_date = (int) implode("",array_reverse(explode("/",$start_date)));
   $end_date = (int) implode("",array_reverse(explode("/",$end_date)));
   if($start_date>$end_date)
   {
     $return = 1; 
   }
   else if($start_date<$end_date)
   {
     $return = 2;  
   }
   return $return;   
}

?>