<?php
class Conexao
{
    private $db;
    private $database_host      = 'mysql03.stillos.com.br';  
    private $database_name      = 'stillos3';
    private $database_user      = 'stillos3';
    private $database_password  = 'h6il_UbzMy!at#';

    private $homologacao_database_host      = '192.168.0.13';  
    private $homologacao_database_name      = 'stillos3';
    private $homologacao_database_user      = 'stillos3';
    private $homologacao_database_password  = 'integracao';        

    public $amb = 'producao';


    // Abre a conex�o com o banco de dados
    function openDataBase()
    {
        if($this->amb == 'producao')
        {
            $this->db = mysql_connect($this->database_host, $this->database_user, $this->database_password);

            if (!$this->db) 
                echo "Erro ao conectar!";
                        
            if (!mysql_select_db($this->database_name))
                echo "Erro ao selecionar o banco de dados!";
        }
        else
        {
            $this->db = mysql_connect($this->homologacao_database_host, $this->homologacao_database_user, $this->homologacao_database_password);

            if (!$this->db)
                echo "Erro ao conectar!";

            if (!mysql_select_db($this->homologacao_database_user))
                echo "Erro ao selecionar o banco de dados!";
        }
    }
    
    // fecha a conex�o com o banco de dados
    function closeDataBase()
    {
    	mysql_close($this->db);
    }
    
    function getLastId() 
    {
    	$result = mssql_fetch_assoc(mssql_query("select @@IDENTITY as id"));
    	return $result['id'];
	}



    ///separar em outra classe
    function convertToISO($str)
    {
        return mb_convert_encoding($str, "ISO-8859-1", mb_detect_encoding($str, "UTF-8, ISO-8859-1, ISO-8859-15", true));
    }

    function getPathDados()
    {
        if($this->amb == 'producao')
            return "e:\\home\\stillos\\dados\\gerencial_foto\\";
        else
            return  "C:\\inetpub\\Dados\\gerencial_foto\\";

    }

}   

?>