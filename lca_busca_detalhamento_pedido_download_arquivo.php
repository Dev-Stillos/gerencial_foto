<?php
// Define o tempo máximo de execução em 0 para as conexões lentas
set_time_limit(0);

$pathFile;
if(isset($_POST['pathFile'])){
	$pathFile = $_POST['pathFile'];
}

if(!file_exists($pathFile)){
	echo "O arquivo não existe.";
	die();
}

$fileName = substr($pathFile, strrpos($pathFile, '\\'));
$fileName = str_replace('\\', "", $fileName);

$fileExtension = substr($pathFile, strrpos($pathFile, '.'));
$tipo;

switch ($fileExtension) {
	case "pdf": $tipo="application/pdf"; break;
	case "exe": $tipo="application/octet-stream"; break;
	case "zip": $tipo="application/zip"; break;
	case "doc": $tipo="application/msword"; break;
	case "xls": $tipo="application/vnd.ms-excel"; break;
	case "ppt": $tipo="application/vnd.ms-powerpoint"; break;
	case "gif": $tipo="image/gif"; break;
	case "png": $tipo="image/png"; break;
	case "jpg": $tipo="image/jpg"; break;
	case "mp3": $tipo="audio/mpeg"; break;
	case "php": // deixar vazio por seurança
	case "htm": // deixar vazio por seurança
	case "html": // deixar vazio por seurança
}




// Configuramos os headers que serão enviados para o browser
header('Content-Description: File Transfer');

if($fileExtension === '.pdf'){
	header('Content-Disposition: inline; filename="'.$fileName.'"');
}
else {
	header('Content-Disposition: attachment; filename="'.$fileName.'"');	
}
header('Content-Type:' . $tipo);
header('Content-Transfer-Encoding: binary');
header('Content-Length: ' . filesize($pathFile));
header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
header('Pragma: public');
header('Expires: 0');

// Envia o arquivo para o cliente
readfile($pathFile);
echo "<script type='text/javascript'>
self.close();
window.close();
window.open('', '_self').close();
window.top.close();
</script>";
die();