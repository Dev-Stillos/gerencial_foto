<?php

$time = microtime(1);
ob_start();
define("_MPDF_TTFONTDATAPATH", 'tempfiles/');
include('includes/Classes_php/MPDF57/mpdf.php');
include('includes/Classes_php/phpmailer/class.phpmailer.php');
require_once 'includes/_connect.php';

function get_content($URL) {
    $ch = curl_init();
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
    curl_setopt($ch, CURLOPT_URL, $URL);
    $data = curl_exec($ch);
    curl_close($ch);
    return $data;
}

function preencheCaracter($str, $caracter_complemento, $qtd) {
	$retorno = $str;

	while (strlen($retorno) < $qtd) {
		$retorno = $caracter_complemento . $retorno;
	}

	return $retorno;
}

function buscaSacado($cod_pedido, $cods_boleto) {
	

	$con = new Conexao();
	$con->openDataBase();


	if ($cod_pedido != '') 
		$sql = "SELECT cod_sacado FROM TB_PEDIDOS_BOLETOS WHERE cod_pedido = " . $cod_pedido;
	else
		$sql = "SELECT cod_sacado FROM TB_PEDIDOS_BOLETOS WHERE cod_boleto in ( " . $cods_boleto . ")";

	$query = mysql_query($sql);
	if ($rs = mysql_fetch_array($query)) 
		$r = $rs['cod_sacado'];

	$con->closeDataBase();

	return $r;
}

$tipo = ''.$_GET['tipo'];
$cod = ''.$_GET['cod'];
$r = array('status' => '', 'msg' => '', 'tempo_execucao' => '');
$arquivos = array();
$cod_sacado = '';
$time = microtime(1);

try {
	//verificar forma de segurança
	$con = new Conexao();
	$con->amb = "producao";

	if ($cod != '') {
		if ($tipo == "boletos" ) {
			$sql = "SELECT * FROM TB_PEDIDOS_BOLETOS WHERE data_pagamento_efetivo is null AND cod_boleto in (".$cod.")";
			$cod_sacado = buscaSacado("", $cod);
		}
		elseif ($tipo == "cod_pedido" ) {
			$sql = "SELECT * FROM TB_PEDIDOS_BOLETOS WHERE data_pagamento_efetivo is null AND cod_pedido = ".$cod;
			$cod_sacado = buscaSacado($cod, "");
		}
		elseif ($tipo == "cod_sacado" ) {
			$sql = "SELECT  *	 FROM TB_PEDIDOS_BOLETOS WHERE data_pagamento_efetivo is null AND cod_sacado = ".$cod;
			$cod_sacado = $cod;
		}
	}

	$con->openDataBase();

	$query = mysql_query($sql);
	while ($rs = mysql_fetch_array($query) ) {

		$resultado = get_content("http://". $_SERVER["HTTP_HOST"] ."/gerencial_foto/cob_impressao_boletos.asp?h=".$rs['hash']);

		$mpdf = new mPDF();
		$mpdf->WriteHTML(utf8_encode($resultado));
		$nome_arquivo = "boleto_".$rs["cod_pedido"]."_".preencheCaracter($rs["parcela"],"0",2).".pdf";
		$mpdf->Output($con->getPathDados()."boletos\\".$nome_arquivo, 'F');		
		unset($mpdf);

		$arquivos[] = $nome_arquivo;
		$cod_sacado = $rs['cod_sacado'];
	}

	$con->closeDataBase();
	$con->openDataBase();

	$sql = "SELECT * FROM TB_SACADOS WHERE cod_sacado = ".$cod_sacado;

	$query = mysql_query($sql);
	//echo $sql;

	if ($rs = mysql_fetch_array($query) ) 
	{

		$from_nome = "Stillo's Reportagens Fotográficas";
		$subject = "Boletos";
		$from  = "noreply@stillos.com.br";
		$host = "smtplw.com.br";
		$username = "stillosadm2";
		$password  = "xwpdkwem4525";
		$mensagem_email = "Prezado(a) [NOME], <br />
							<br />
							Segue anexo os boletos referentes a sua compra na Stillos. 
							<br /><br /><br />
                            Atenciosamente<br />
                            Stillos Reportagens Fotográficas<br /><br /><br />
							*****************************************************<br />
							Envio automatico, por favor não responda este e-mail.<br />
							*****************************************************<br />";

		$mensagem_email = str_replace("[NOME]", $rs['nome'], $mensagem_email);

		$mail = new PHPMailer();
		$mail->IsSMTP();
		$mail->SMTPAuth = true;
		$mail->Host = $host;
		$mail->Port = '587';
		$mail->Username = $username;
		$mail->Password = $password;
		$mail->From = $from;
		$mail->FromName = utf8_decode($from_nome);
		$mail->IsHTML(true);

		$mail->Subject = utf8_decode($subject);
		$mail->Body = utf8_decode($mensagem_email);

		$mail->AddAddress($rs['email']);
        // $mail->AddAddress("marcelo@stillos.com.br");
		//$mail->AddBCC("marcelo.beltrame@gmail.com");
		
		foreach ($arquivos as $arquivo) {
			echo $con->getPathDados()."boletos\\".$arquivo.'<br>';
			$mail->AddAttachment($con->getPathDados()."boletos\\".$arquivo);
		}
		$mail->Send();
	}

	$r['status'] = '0';
	$r['msg'] = 'boletos enviados com sucesso!';

	$con->closeDataBase();
} 
catch (Exception $e) {
    echo "The exception was thrown on line: " . $e->getLine();
    $r['status'] = $e->getCode();
	$r['msg'] =  $e->getMessage();
}

$r['tempo_execucao'] =  ((microtime(1) - $time));
ob_end_clean();
echo json_encode($r);

?>