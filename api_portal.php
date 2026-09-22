<?php
/**
 * API PORTAL STILLOS - v1.2 (Paginada e Refinada)
 * Compatível com versões antigas de PHP (IIS/Locaweb).
 */

error_reporting(0);
header('Content-Type: application/json; charset=utf-8');

// Configurações de Segurança
$MY_API_KEY = "Stillos_Full_Access_2026";

$receivedKey = "";
if (isset($_SERVER['HTTP_X_API_KEY'])) {
    $receivedKey = $_SERVER['HTTP_X_API_KEY'];
} elseif (isset($_GET['key'])) {
    $receivedKey = $_GET['key'];
}

if ($receivedKey !== $MY_API_KEY) {
    header('HTTP/1.1 403 Forbidden');
    echo json_encode(array("error" => "Acesso negado"));
    exit;
}

// Configurações do Banco de Dados
$db_host = "mysql03.stillos.com.br";
$db_user = "stillos3";
$db_pass = "h6il_UbzMy!at#";
$db_name = "stillos3";

$conn = mysqli_connect($db_host, $db_user, $db_pass, $db_name);
mysqli_set_charset($conn, "utf8");

if (!$conn) {
    echo json_encode(array("error" => "Falha na conexão com o banco"));
    exit;
}

$action = isset($_GET['action']) ? $_GET['action'] : '';

// --- AÇÃO: LISTAR PRODUÇÕES (REFINADA) ---
if ($action == 'get_productions') {
    $busca_nome = isset($_GET['busca_nome']) ? $_GET['busca_nome'] : '';
    $busca_id = isset($_GET['busca_codigo']) ? $_GET['busca_codigo'] : '';
    $busca_contrato = isset($_GET['busca_contrato']) ? $_GET['busca_contrato'] : '';
    $filtro = isset($_GET['filtro']) ? (int)$_GET['filtro'] : 2; 
    $pagina = isset($_GET['pagina']) ? (int)$_GET['pagina'] : 1;
    $limit = 10; // Máximo 10 linhas conforme pedido
    $offset = ($pagina - 1) * $limit;

    $where = " WHERE 1=1 ";
    if ($filtro == 1) $where .= " AND PRODUCAO_FINALIZADA = 1 ";
    if ($filtro == 2) $where .= " AND PRODUCAO_FINALIZADA = 2 ";
    if ($filtro == 3) $where .= " AND EXISTS (SELECT 1 FROM TB_FICHAS_IDENTIFICACAO f WHERE f.COD_PRODUCAO = p.COD_PRODUCAO AND f.ALBUM_DIGITAL = 1) ";
    
    if ($busca_nome) $where .= " AND DESCRICAO LIKE '%".mysqli_real_escape_string($conn, $busca_nome)."%' ";
    if ($busca_id) $where .= " AND COD_PRODUCAO = '".mysqli_real_escape_string($conn, $busca_id)."' ";
    if ($busca_contrato) {
        $where .= " AND EXISTS (SELECT 1 FROM TB_FICHAS_IDENTIFICACAO f2 WHERE f2.COD_PRODUCAO = p.COD_PRODUCAO AND f2.COD_CONTRATO LIKE '%".mysqli_real_escape_string($conn, $busca_contrato)."%') ";
    }

    // Contagem Total Filtrada
    $sqlFiltered = "SELECT COUNT(*) FROM TB_PRODUCAO p $where";
    $resFiltered = mysqli_query($conn, $sqlFiltered);
    $totalFiltered = 0;
    if ($resFiltered) {
        $rowF = mysqli_fetch_row($resFiltered);
        $totalFiltered = (int)$rowF[0];
    }

    // Contagem Total Global (Independente de Filtros)
    $sqlGlobal = "SELECT COUNT(*) FROM TB_PRODUCAO";
    $resGlobal = mysqli_query($conn, $sqlGlobal);
    $totalGlobal = 0;
    if ($resGlobal) {
        $rowG = mysqli_fetch_row($resGlobal);
        $totalGlobal = (int)$rowG[0];
    }

    // Query de Dados
    $sql = "SELECT COD_PRODUCAO, DESCRICAO, DATA_FINALIZADA, PRODUCAO_FINALIZADA, PRODUTOS_AGREDADOS FROM TB_PRODUCAO p $where ORDER BY COD_PRODUCAO DESC LIMIT $offset, $limit";
    $result = mysqli_query($conn, $sql);
    
    $productions = array();
    while ($row = mysqli_fetch_assoc($result)) {
        $id = $row['COD_PRODUCAO'];
        
        // Contratos
        $resSub = mysqli_query($conn, "SELECT DISTINCT(COD_CONTRATO) FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = '$id'");
        $contratos = array();
        while($s = mysqli_fetch_row($resSub)) $contratos[] = $s[0];
        
        // Lotes
        $resSub = mysqli_query($conn, "SELECT COUNT(*) FROM TB_PRODUCAO_LOTES WHERE COD_PRODUCAO = '$id'");
        $qtdLotes = 0;
        if ($resSub) {
            $rowL = mysqli_fetch_row($resSub);
            $qtdLotes = $rowL[0];
        }
        
        // Estatísticas e Pastas (Mín/Máx)
        $resSub = mysqli_query($conn, "SELECT COUNT(*) as qtd, SUM(INICIAL) as fotos, MAX(ALBUM_DIGITAL) as digital, MIN(numero_pasta_gerencia) as min_p, MAX(numero_pasta_gerencia) as max_p FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = '$id'");
        $stats = mysqli_fetch_assoc($resSub);

        $productions[] = array(
            "id" => $id,
            "nome" => $row['DESCRICAO'],
            "contratos" => implode(", ", $contratos),
            "dataFinalizada" => $row['DATA_FINALIZADA'] ? $row['DATA_FINALIZADA'] : "",
            "produtosAgregados" => $row['PRODUTOS_AGREDADOS'],
            "qtdLotes" => (int)$qtdLotes,
            "qtdAlbuns" => (int)$stats['qtd'],
            "minPasta" => $stats['min_p'],
            "maxPasta" => $stats['max_p'],
            "totalFotos" => (int)$stats['fotos'],
            "isDigital" => ($stats['digital'] == 1),
            "originalStatus" => (int)$row['PRODUCAO_FINALIZADA']
        );
    }
    
    // Retorno com Metadados de Paginação
    echo json_encode(array(
        "data" => $productions,
        "total" => $totalFiltered,
        "total_global" => $totalGlobal,
        "page" => $pagina,
        "limit" => $limit,
        "total_pages" => ceil($totalFiltered / $limit)
    ));
}

// --- AÇÃO: LISTAR FORMANDOS ---
elseif ($action == 'get_students') {
    $cod_prod = isset($_GET['cod_prod']) ? (int)$_GET['cod_prod'] : 0;
    
    if ($cod_prod == 0) {
        echo json_encode(array("error" => "ID Inválido"));
        exit;
    }

    $sql = "SELECT * FROM TB_FICHAS_IDENTIFICACAO WHERE COD_PRODUCAO = $cod_prod ORDER BY NOME_FORMANDO ASC";
    $result = mysqli_query($conn, $sql);
    
    $students = array();
    while ($row = mysqli_fetch_assoc($result)) {
        $students[] = array(
            "id" => $row['COD_FORMANDO'],
            "nome" => $row['NOME_FORMANDO'],
            "contrato" => $row['COD_CONTRATO'],
            "pasta" => $row['numero_pasta_gerencia'],
            "albumFisico" => $row['ALBUM'],
            "albumDigital" => (int)$row['ALBUM_DIGITAL'],
            "inicial" => (int)$row['INICIAL'],
            "comissao" => (int)$row['COMISSAO'],
            "brinde" => (int)$row['BRINDE'],
            "prodAgregado" => (int)$row['PRODUTO_AGREGADO'],
            "lote" => $row['LOTE']
        );
    }
    echo json_encode($students);
}

else {
    echo json_encode(array("status" => "Portal API Ativa", "server_time" => date('Y-m-d H:i:s')));
}

mysqli_close($conn);
?>
