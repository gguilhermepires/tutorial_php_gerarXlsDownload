<?php

	/**
	 * Relatorio p;
	 *
	 * @package Relatorio
	 * @author Guilherme Pires
	 *
	 * @return (file) um arquivo em xls ;
	 */

	$fileName = "nomearq" . "_" . date_format(new DateTime("Now"), 'd-m-Y').".xls";

	// header('Content-Type: application/vnd.ms-excel');
	header('Content-Type: application/excel');
    header('Content-Disposition: attachment;filename="'.$fileName.'"');
    header('Cache-Control: max-age=0');

/**
 * @define (string) DS - Directory separator.
 */
define("DS", "/", true);

$base_path = "";
for ($i = 2; $i < (int) substr_count($_SERVER['PHP_SELF'], DS); $i++) {
	$base_path .= '../';
}

/**
 * @define (string) BASE_PATH_OPEN - Best Root Path.
 */
define("BASE_PATH_OPEN", $base_path, true);

/** Files needed for the application **/
include_once BASE_PATH_OPEN . 'util.class.php';


include_once BASE_PATH_OPEN . 'dao' . DS . 'conexaoBanco.class.php';

$con = new Conexao();

/***************************************************************************************************************
Variaveis  - passadas p/ o calculo;
/**************************************************************************************************************/

$id = $_SESSION['...'];

$con = $con->getConexao(); 



$query = "
	SELECT **";

$stmt = $con->query($query);
$resultado_tc = $stmt->fetch();

/****************************************************************************************************************
- -   IMPRIMINDO xls   - -
/***************************************************************************************************************/

$query = "
    SELECT *... ";
$stmt = $con->query($query);
$resultado = $stmt->fetchAll();

$dadosXls  = "";

$dadosXls .= "<h5>".utf8_decode("Relatório  ....")."</h5>";
$dadosXls .= "<hr>";
$dadosXls .="<p>".utf8_decode("Por Cliente")."</p>";
$dadosXls .= "<p>".utf8_decode("Referência: " . $resultado["..."]."</p>");
$dadosXls .= "<p>".utf8_decode("Data de Emissão: " . date_format(new DateTime("Now"), 'd/m/Y  à\s  H\hi')."</p>");
$dadosXls .= " <hr>";
$dadosXls .= "  <table border='1' >";
$dadosXls .= "          <tr>";
$dadosXls .= "          <th>".utf8_decode("- CODIGO -")."</th>";
$dadosXls .= "          <th>".utf8_decode("- PRODUTO -")."</th>";
$dadosXls .= "          <th>".utf8_decode("- MARCA -")."</th>";
$dadosXls .= "          <th>".utf8_decode("- VALOR -")."</th>";
$dadosXls .= "      </tr>";
	
foreach ($resultado as $dadosPedido) {

    
   
	$dadosXls .= "      <tr>";
	$dadosXls .= "          <td>".utf8_decode(str_pad($dadosPedido['..'], 15, ' ', STR_PAD_LEFT))."</td>";
	$dadosXls .= "          <td>".utf8_decode($dadosPedido['..'])."</td>";
	$dadosXls .= "          <td>".utf8_decode("  " . str_pad($dadosPedido['..'], 52, ' ', STR_PAD_BOTH))."</td>";
	$dadosXls .= "          <td>".str_pad(Util::mostraValor($novo_valor, 2), 28, ' ', STR_PAD_LEFT)."</td>";
	
	$dadosXls .= "      </tr>";
}

$dadosXls .= "  </table>";

echo $dadosXls;
?>
