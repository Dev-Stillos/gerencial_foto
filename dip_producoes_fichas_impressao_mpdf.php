 	<?php
	if ($_SERVER["REQUEST_METHOD"] != "POST") {
		echo "Você precisa acessar esse link direto pelo botão da produção"."<br />"."Por favor, acesse o sistema da foto e tente novamente.";
		die();
	}

	//Aqui eu checo se houve um reload na página
	$pageWasRefreshed = isset($_SERVER['HTTP_CACHE_CONTROL']) && $_SERVER['HTTP_CACHE_CONTROL'] === 'max-age=0';
	if (!$pageWasRefreshed) {
		echo "Não é permitido recarregar a página, pois assim previnimos que não haja erros no relatório"."<br />"."Por favor, acesse o sistema da foto e tente novamente.";
		die();
	}

	//VERIFICO SE FOI ENVIADO OS FORMANDOS, CO E FOTOS
	if(!isset($_POST["formandos"]) || !isset($_POST["co"]) || !isset($_POST["fotos"]) || !isset($_POST["nome"]) ){
		echo "Parâmetros não enviados, por favor, acesse o sistema da foto e tente novamente.";
		die();
	}

	$formandos      = $_POST["formandos"];
	$co             = $_POST["co"];
	$fotos          = $_POST["fotos"];
	$nome           = $_POST["nome"];
	$totalFotos     = $_POST["total_fotos"];
	$totalFormandos = $_POST["total_formandos"];


	$dados = json_encode(array(
		"formandos"      => $formandos,
		"co"             => $co,
		"fotos"          => $fotos,
		"nome"           => $nome,
		"totalFotos"     => $totalFotos,
		"totalFormandos" => $totalFormandos
	));

	require_once __DIR__ . './vendor/autoload.php';

	$url= 'http://".$_SERVER["HTTP_HOST"]."/gerencial_foto/dip_producoes_fichas_impressao_layout.php';
 	//$urlBoleto= 'http://192.168.0.197:3030//formandos/boleto/gera_boleto_pdf.php';

	$ch = curl_init();
	$timeout = 5;
	curl_setopt($ch, CURLOPT_URL, $url);
	curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
	curl_setopt($ch, CURLOPT_POSTFIELDS, $dados);
	curl_setopt($ch, CURLOPT_POST, 1);
	curl_setopt($ch, CURLOPT_CONNECTTIMEOUT, $timeout);
	$resultado = curl_exec($ch);
	curl_close($ch);



	$mpdf = new \Mpdf\Mpdf();
	$mpdf->allow_charset_conversion=true;
	$mpdf->charset_in='UTF-8';
	$csspdf = '.A4 {
			background: white;
			width: 21cm;
			height: 29.7cm;
			display: block;
			margin: 0 auto;
			padding: 10px 25px;
			margin-bottom: 0.5cm;
			box-shadow: 0 0 0.5cm rgba(0, 0, 0, 0.5);
			overflow-y: scroll;
			box-sizing: border-box;
			font-size: 12pt;
	}

	div#tabela {
			margin: 15mm auto;
			display: block;
	}

	table#tabela {
			display: inline-block;
			padding: 10px;
			margin: 5px auto;

			position: relative;

 		 	left: 50%;
 		 	transform: translateX(-50%);
	}


	/* TABELA STYLE INTERNET*/
	.responstable {
			margin: 1em 0;
			width: 100%;
			overflow: hidden;
			background: #FFF;
			color: #151515;
			border-radius: 10px;
			border: 1px solid #d3d3d3;
	}

	.responstable tr {
			border: 1px solid #D9E4E6;
	}

	.responstable tr:nth-child(odd) {
			background-color: #cccccc;
	}

	.responstable th {
			border: 1px solid #FFF;
			background-color: #3f3f3f;
			color: #FFF;
			padding: 1em;
			margin-right: 8mm;
	}

	.responstable th:last-child {
			margin-right: 0;
	}

	.responstable th:first-child {
			display: table-cell;
			text-align: center;
	}

	.responstable th:nth-child(2) {
			display: table-cell;
	}

	.responstable th:nth-child(2) span {
			display: none;
	}

	.responstable th:nth-child(2):after {
			content: attr(data-th);
	}

	.responstable td {
			word-wrap: break-word;
			max-width: 7em;
	}

	.responstable td:first-child {
			display: table-cell;
			text-align: center;
			border-right: 1px solid #e0e0e0;
	}

	.responstable tr:nth-child(odd) td:first-child {
			border-right: 1px solid #f1f1f1;
	}

	.responstable th,
	.responstable td {
			text-align: center;
			margin: .5em 1em;
			/*margin-right: 4em;*/
	}

	/* TABELA STYLE INTERNET*/

	#caption1 {
			margin-bottom: 10px;
	}

	@media print {
			.page-break {
					display: block;
					page-break-before: always;
			}
			size: A4 portrait;
	}

	@media print {
			@page {
					size: auto;
					@top-left {
							content: "";
					}
			}
			* {
					background: transparent !important;
					color: #000 !important;
					text-shadow: none !important;
					filter: none !important;
					-ms-filter: none !important;
			}
			body {
					margin: 0;
					padding: 0;
			}
			a[href]:after {
					content: none !important;
					text-decoration: none;
					color: black;
					cursor: default;
			}
			a:after,
			a[href^="javascript:"]:after,
			a[href^="#"]:after {
					content: "" !important;
			}
			nav,
			footer,
			video,
			audio,
			object,
			embed {
					display: none !important;
			}
			.A4 {
					box-shadow: none;
					margin: 0;
					width: auto;
					height: auto;
					overflow-y: hidden;
			}
			div#tabela {
					margin: 0;
			}
			.noprint {
					display: none;
			}
			.enable-print {
					display: block;
			}
			.a4-display-none {
					display: none;
			}

			/*ARRUMANDO A TABELA*/
			.responstable {
					border-radius: 0;
					border: none;
					color: #000;
			}
			.responstable tr:nth-child(odd) {
					background-color: #CCC;
					color: #000;
			}
			.responstable th {
					border: 1px solid #cecece;
					background-color: #000;
					color: #FFF;
			}
	}
';

	// $css =  get_included_files('style.css')
	$mpdf->WriteHTML($csspdf,1);
	$mpdf->WriteHTML($resultado, 2);
	$mpdf->Output();




?>

<?php
	$data = json_decode(file_get_contents('php://input'), true);
	
	$formandos      = $data["formandos"];
	$co             = $data["co"];
	$fotos          = $data["fotos"];
	$nome           = $data["nome"];
	$totalFotos     = $data["totalFotos"];
	$totalFormandos = $data["totalFormandos"];
	// $url= 'http://".$_SERVER["HTTP_HOST"]."/gerencial_foto/dip_producoes_fichas_impressao_mpdf.php';
	//
	// $ch =  curl_init(); // Inicia curl
	// curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
	// curl_setopt($ch, CURLOPT_RETURNTRANSFER, true); // PEGAR OS DADOS
	// curl_setopt($ch, CURLOPT_URL, $url); // Executa a url
	// $resultado = curl_exec($ch);//Executar todo curl
	//
	// curl_close($ch); // Fecha curl
	//
	// $returnJson = json_decode($resultado, true);
	//
	// $formandos      = $returnJson->{'formandos'};
	// $co             = $returnJson->{'co'};
	// $fotos          = $returnJson->{'fotos'};
	// $nome           = $returnJson->{'nome'};
	// $totalFotos     = $returnJson->{'total_fotos'};
	// $totalFormandos = $returnJson->{'total_formandos'};
?>
<html lang="pt-br">
<head>
	<meta charset="utf-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<link rel="stylesheet" href="./css/style_dip_producoes_fichas_impressao_layout.css">
</head>
<body>
	<div class=A4>
		<h1 class=a4-display-none text-center>
			Impressão Resumo da divisão por lotes
		</h1>
		<div id=tabela>
			<table id=tabela class=responstable>
				<caption id="caption1"><?php echo  $nome ?></caption>
				<caption>Resumo da divisão por lotes</caption>
				<thead>
					<tr>
						<th>Lote</th>
						<th>Qtd. de Formandos</th>
						<th>Qtd. de Co</th>
						<th>Qtd. de Fotos</th>
					</tr>
				</thead>
				<tbody>
					<?php for ($i = 0; $i < count($formandos); $i++) {?>
					<tr>
						<td><?php echo ($i + 1) ?></td>
						<td><?php echo $formandos[$i] ?></td>
						<td><?php echo $co[$i] ?></td>
						<td><?php echo $fotos[$i] ?></td>
					</tr>
					<?php } ?>
				</tbody>
				<tfoot>
					<tr>
						<td colspan="2" style="padding: 8px;">
						Total de Formandos: <span style="text-align: right; font-weight: bold; "><?php echo $totalFormandos ?></span>
						</td>
						<td colspan="2" style="padding: 8px;">
						Total de Fotos: <span style="text-align: right; font-weight: bold;"><?php echo $totalFotos ?></span>
						</td>
					</tr>
				</tfoot>
			</table>
		</div>
	</div>
</body>
</html>
