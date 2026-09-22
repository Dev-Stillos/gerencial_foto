<?php

echo buscaCep($_POST['cep']);

function buscaCep($cep)
{
	$cep_limpo = str_replace('-', '', $cep);
	$cep_mascara = substr ( $cep_limpo, 0, 5) . '-' . substr ( $cep_limpo, 5, 3);
	



	$retorno = cepAberto($cep_limpo);

	if (!isset($retorno) || $retorno == '' || $retorno == '{}'|| $retorno == '[]' )
	{
		//$retorno = cepLivre($cep_mascara);
	}

	if (!isset($retorno) || $retorno == '' || $retorno == '{}'|| $retorno == '[]' )
	{
		$retorno = '';
	}

	return $retorno;
}

function cepLivre($cep)
{
	$key = 'c324a9844ab83c93f7c04a56c55c504b';
libxml_use_internal_errors(true) ;


	$out = json_encode(simplexml_load_file("http://ceplivre.com.br/consultar/cep/$key/$cep/xml"));
	$out = json_decode($out);
	$retorno =  array();

	if(isset($out->cep->logradouro))
	{
		$retorno['logradouro'] =  $out->cep->tp_logradouro .' '. $out->cep->logradouro;
		$retorno['cep'] = $out->cep->cep;
		$retorno['cidade'] = $out->cep->cidade;
		$retorno['estado'] = $out->cep->uf_sigla;
		$retorno['bairro'] = $out->cep->bairro;
	}

return json_encode($retorno);
}

function cepAberto($cep)
{
	$key = '9187d031c851fd050f9d5562a595611b';
	$url = 'http://www.cepaberto.com/api/v2/ceps.json?cep='.$cep;
	$ch = curl_init();
	curl_setopt($ch, CURLOPT_URL, $url );
	curl_setopt($ch, CURLOPT_HTTPHEADER, array('Authorization: Token token="'.$key.'"'));
	curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
	$out = curl_exec($ch);

	$i = 0;

	//Se houver erro por muitas tentativas 
	//limitamos a 5 tentativas 
	while (trim($out) == '403 Forbidden (Rate Limit Exceeded)' && $i < 5 ) { 
		sleep(5);
		$out = curl_exec($ch);
		$i ++;
	}
	
	$retorno = array();
	//se não houver erro então convertemos para nosso formato
	if(!strpos($out , "403") && !strpos($out,"404") && strlen(trim($out)) > 0 && $out !== '{}')
	{
		$out = json_decode($out);	
		$retorno['logradouro'] = $out->logradouro;
		$retorno['cep'] = $out->cep;
		$retorno['cidade'] = $out->cidade;
		$retorno['estado'] = $out->estado;
		$retorno['bairro'] = $out->bairro;
	}

return json_encode($retorno);
}
?>