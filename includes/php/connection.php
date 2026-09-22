<?php

function openConnection(){
   $mysql_connection = mysql_connect("192.168.0.30","stillos","stillos") or die ('Erro ao conectar na base de dados');     
   mysql_selectdb("gerencial_foto",$mysql_connection) or die ('Erro ao selecionar o banco de dados');
}

?>