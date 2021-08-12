<?php

require_once __DIR__ . '/report.php';

$loja = str_replace('"','',$_POST['lojas']);
$loja2 = str_replace(array('"', "'", ' ', ','), '_', $loja);
$mes = isset($_POST['mes']) ? $_POST['mes'] : null;
$filename = implode(' - ', array_filter([$loja2, $mes]));

$xls = new PHPExcel();

$styleArray = [
    'font' => [
        'size' => 10
    ],
    'alignment' => [
        'vertical' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
    ]

];



$activeSheetIndex = 0;
$xls->setActiveSheetIndex(0);

$sheet = $xls->getActiveSheet();
$sheet->setShowGridlines(false);

$sheet->setTitle('Cadastro');

$records = detalhesContrato($sheet, $loja);
$records = aluguelContratual($sheet, $loja, $records);
$records = percentualVenda($sheet, $loja, $records);
reducaoAluguel($sheet, $loja, $records++);




if ($mes != null) {
    $sheet = $xls->createSheet(2);
    $sheet->setTitle('R$');
    $sheet->setShowGridlines(false);
    $records = 1;
    $records = valorMensal($sheet, $loja, $mes, 'performanceabs', $records);
    $records = visaoMeses($sheet, $loja, $mes, 'performanceabs', $records);
    shoppings($sheet, $loja, $mes, 'performanceabs', $records);

    $sheet = $xls->createSheet(3);
    $sheet->setTitle('R$ por m²');
    $sheet->setShowGridlines(false);
    $records = 1;
    $records = valorMensal($sheet, $loja, $mes, 'performancem2', $records);
    $records = visaoMeses($sheet, $loja, $mes, 'performancem2', $records);
    shoppings($sheet, $loja, $mes, 'performancem2', $records);


}

$xls->getDefaultStyle()->applyFromArray($styleArray);
foreach (range('A', $xls->getActiveSheet()->getHighestDataColumn()) as $col) {
                 $xls->getActiveSheet()
                ->getColumnDimension($col)
                ->setAutoSize(false)
                ->setWidth('19');
    }


$xls->getActiveSheet()->getStyle('H105:H'.$xls->getActiveSheet()->getHighestRow())
    ->getAlignment()->setWrapText(true);

header("Expires: Mon, 1 Apr 1974 05:00:00 GMT");
header("Last-Modified: " . gmdate("D,d M YH:i:s") . " GMT");
header("Cache-Control: no-cache, must-revalidate");
header("Pragma: no-cache");
header("Content-type: application/vnd.ms-excel");
header("Content-Disposition: attachment; filename={$filename}.xls");
//header('Content-Disposition: attachment; filename= str_replace(array('"', "'", ' ', ','), '_', $loja2.xls');

$objWriter = new PHPExcel_Writer_Excel5($xls);
$objWriter->save('php://output');
