<?php
define('PATH', getcwd());
define('PSEPARATOR', DIRECTORY_SEPARATOR);

ini_set('allow_url_fopen',1);
ini_set('allow_url_include',1);
ini_set('memory_limit','256M'); 

//error_reporting(E_ERROR | E_PARSE);

function construtor()
{
	if(!isset($_GET["colacao"]) || !isset($_GET["temporada"]))
	{
		echo "Erro, favor selecionar uma colação e uma temporada";
		die();
	}
	if(empty($_GET["colacao"]))
	{
		echo "Erro, favor selecionar uma colação";
		die();
	}
	if(empty($_GET["temporada"]))
	{
		echo "Erro, favor selecionar uma temporada";
		die();
	}	
}


function getColacao()
{
	if(isset($_GET["colacao"]))
		return filter_var($_GET["colacao"],FILTER_SANITIZE_STRING);
}

function getTemporada()
{
	if(isset($_GET["temporada"]))
		return filter_var($_GET["temporada"],FILTER_SANITIZE_STRING);
}

function getImagesDirectory()
{
	return PATH . PSEPARATOR . '..'. PSEPARATOR . '..' . PSEPARATOR .'Dados' . PSEPARATOR . 'gerencial_foto' . PSEPARATOR . 'mck_identificacao'. PSEPARATOR;
}

function getConexao()
{
	$dsn = 'mysql:host=mysql03.stillos.com.br;port=3306;dbname=stillos3;charset=utf8';
	$user = 'stillos3';
	$password = 'hauryecO@ys';
	$options = array(
		PDO::ATTR_PERSISTENT => true
		);

	return new PDO($dsn, $user, $password, $options);
}


function executeQuery($query)
{
	try
	{

		$sql = getConexao()->prepare($query);
		$sql->execute();
		$dados = array();
		
		if($sql)
		{
			while($row = $sql->fetch(PDO::FETCH_ASSOC, PDO::FETCH_ORI_NEXT))
			{
				$nome_arquivo = $row["nome_arquivo"];
				/*$dados[] = array(
					"colacao" => $row["cod_mack_colacao"],
					"formando" => $row["cod_mack_formando"],
					"nome" => $row["nome"],
					"tia" => $row["tia"],
					"nome_arquivo" => $row["nome_arquivo"]
				);*/
				if($nome_arquivo)
				{
					array_push($dados, $nome_arquivo);
				}
			}
		}
		return $dados;
	}
	catch(Exception $e)
	{
		echo "Erro: " . $e->getMessage();
		die();
	}
}

function getCodMackFormando()
{
	$query = "SELECT cod_mack_colacao, cod_mack_formando, nome, tia, nome_arquivo FROM VW_TB_MACK_FORMANDOS";
	$query .= " WHERE cod_mack_temporada = ";
	$query .= getTemporada();
	$query .= " AND cod_mack_colacao = ";
	$query .= getColacao();
	$query .= " ORDER BY cod_mack_formando";

	return executeQuery($query);
}

function getImagesName()
{
	return getCodMackFormando();
}

function listImages()
{
	$files  = scandir(getImagesDirectory());
	return $files;
}

function myFileGetContents( $site_url ){
	$ch = curl_init();
	$timeout = 5; // set to zero for no timeout
	
	curl_setopt ($ch, CURLOPT_URL, $site_url);
	curl_setopt ($ch, CURLOPT_CONNECTTIMEOUT, $timeout);

	ob_start();
	curl_exec($ch);

	curl_close($ch);
	$file_contents = ob_get_contents();
	
	ob_end_clean();
	
	return $file_contents;
}

function zipDownload($file_names,$archive_file_name,$file_path) {
	 /* Create object */
	 $zip = new ZipArchive();
	 
	 /* Create file or error if unsuccessful */
	 if ($zip->open($archive_file_name, ZIPARCHIVE::CREATE )!==TRUE) {
	  exit("cannot open <$archive_file_name>\n");
	 }
	 
	 if(empty($file_names) || $file_names == null)
	 {
	 	echo "Essa colação não possui nenhum formando com foto cadastrada";
	 	die();
	 }

	  /* Add each file $file_name array to archive */
	  foreach($file_names as $files)
	  {
	   	$zip->addFile($file_path.$files,$files);
	  }
	 
	 $zip->close();

	/* Send headers to force download */
	header("Content-type: application/zip");
	header("Content-Disposition: attachment; filename=$archive_file_name");
	header('Content-Length: ' . filesize($file));
	header("Pragma: no-cache");
	header("Expires: 0");
	ob_clean();
	if(file_exists($archive_file_name))
	{
		readfile($archive_file_name);
	}
	else
	{
		echo "Não foi possível criar o arquivo";
		die();
	} 
	 
	/* Comment out this line if you don't want to delete file after download */
	@unlink("$archive_file_name");
	exit;
}

construtor();
 
$file_names = getImagesName();
$archive_file_name = "formandos_". getColacao(). getTemporada() .".zip";
 
zipDownload($file_names,$archive_file_name, getImagesDirectory());

die();

 ?>

