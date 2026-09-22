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
        <caption id="caption1"><?php echo  $nome; ?> <br>Resumo da divisão por lotes </caption>
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
