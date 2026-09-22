<?php
/** Error reporting */
set_time_limit(3600);
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('America/Sao_Paulo');
define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');
header('Content-type: text/html; charset=ISO-8859-1');

/** Include PHPExcel e Conexao*/
require_once 'includes/_connect.php';
require_once 'includes/Classes_php/PHPExcel/IOFactory.php';

$con = new Conexao();


//Recebemos o codigo da lista a ser importada
$cod_mack_lista_provaveis_formandos = $_GET['l'];

//index planilha atual
if ( isset($_GET['idx']) )
    $index_atual = $_GET['idx'];
else
    $index_atual = 0;

$con->openDataBase();
//Buscamos a lista
$sql = mysql_query("SELECT * FROM TB_MACK_LISTAS_PROVAVEIS_FORMANDOS WHERE status = 0 AND cod_mack_lista_provaveis_formandos = ".$cod_mack_lista_provaveis_formandos);
if($rs = mysql_fetch_assoc($sql))
    $inputFileName = $rs['nome_arquivo_servidor'].'.'.$rs['extencao'];
else
    die("Arquivo já importado");

$con->closeDataBase();

//Read your Excel workbook
try {
    $inputFileType = PHPExcel_IOFactory::identify('../../Dados/gerencial_foto/mck_listas/'.$inputFileName);
    $objReader = PHPExcel_IOFactory::createReader($inputFileType);
    $objPHPExcel = $objReader->load('../../Dados/gerencial_foto/mck_listas/'.$inputFileName);
} catch(Exception $e) {
    die('Erro ao carregar o arquivo "'.pathinfo($inputFileName,PATHINFO_BASENAME).'": '.$e->getMessage());
}
$i = 0 ;

//Buscamos quanitdade de planilhas
$qtd_planilhas = $objPHPExcel->getSheetCount();

if($index_atual < $qtd_planilhas)
{
    $con->openDataBase();
    $planilha = $objPHPExcel->getSheet($index_atual);

    $departamento_sigla = $planilha->getTitle(); 
    $highestColumn = 'C';
    $highestRow = $planilha->getHighestRow(); 

    // Loop para percorrermos as linhas da planilha
    for ($row = 1; $row <= $highestRow; $row++){ 
        // resgatamos valor da linha
        $rowData = $planilha->rangeToArray('A' . $row . ':' . $highestColumn . $row,
                                        NULL,
                                        FALSE,
                                        FALSE);

        //Somente importamos se ambos os campos forem diferentes de ''
        if ($rowData[0][1].'' != '' && $rowData[0][2].'' != '') 
        {
            //Veificamos se a linha não se trata de um cabecalho
            if ($rowData[0][1] != 'MATRICULA')
                if ($departamento_sigla == $rowData[0][1]) //Atualizamos o curso
                {
                   $curso = $con->convertToISO($rowData[0][2]);
                }
                else //importamos
                {
                    $matricula = $con->convertToISO($rowData[0][1]);
                    $nome = $con->convertToISO($rowData[0][2]);

                    //insermos
                    $sql = "INSERT INTO TB_MACK_LISTAS_PROVAVEIS_FORMANDOS_NOMES (
                                cod_mack_lista_provaveis_formandos,
                                nome,
                                matricula,
                                curso,
                                departamento,
                                departamento_sigla
                                ) VALUE (
                                $cod_mack_lista_provaveis_formandos,
                                '$nome',
                                '$matricula',
                                '$curso',
                                '$departamento',
                                '$departamento_sigla'
                            )";

                    mysql_query($sql);
                }
        }
        if ($rowData[0][1].'' == '' && $rowData[0][2].'' != '')
        {
            //nome do departamento
            $departamento = $con->convertToISO($rowData[0][2]);
        }
    }
    $index_atual ++;

    header("Location: mck_lista_provaveis_formandos_formulario_processar.php?idx=".$index_atual."&l=".$cod_mack_lista_provaveis_formandos);
    $con->closeDataBase();
}
else
{
    $con->openDataBase();
    //Atualizamos o status
    $sql = "UPDATE TB_MACK_LISTAS_PROVAVEIS_FORMANDOS SET status = 1 WHERE cod_mack_lista_provaveis_formandos = ".$cod_mack_lista_provaveis_formandos;

    mysql_query($sql);

    $con->closeDataBase();

    header("Location: mck_lista_provaveis_formandos.asp?msg=1");
}


?>