<?php
require_once __DIR__ . '/../config.php';
require_once __DIR__ . '/../functions.php';

require_once __DIR__ . '/../lib/Classes/PHPExcel.php';
require_once __DIR__ . '/../lib/Classes/PHPExcel/Cell.php';

class Formatter
{
    public static function moeda($value)
    {
        return moeda($value);
    }

    public static function moedaIfNumeric($value)
    {
        if (is_numeric($value)) {
            return moeda($value);
        }

        return $value;
    }

    public static function verificaNuloIfNumeric($value)
    {
        if (is_numeric($value)) {
            return verifica_nulo(moeda($value));
        }

        return verifica_nulo($value);
    }
}

/**
 * @param PHPExcel_Worksheet $sheet
 * @param int $columnsCount
 * @param string $title
 * @param int $row
 * @return void
 */
function title($sheet, $columnsCount, $title, $row = 1)
{
    $sheet->getRowDimension($row)->setRowHeight(24);
    $sheet->setCellValueByColumnAndRow(0, $row, $title);
    $style = $sheet->getStyleByColumnAndRow(0, $row, $columnsCount - 1, $row);
    $style->getFont()->setBold(true);
    $style->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('F7F7F7');
}

/**
 * @param PHPExcel_Worksheet $sheet
 * @param $columns
 * @param int $row
 * @return void
 */
function columns($sheet, $columns, $row = 2)
{
    $sheet->getRowDimension($row)->setRowHeight(20);
    foreach ($columns as $index => $column) {
        $cellId = PHPExcel_Cell::stringFromColumnIndex($index) . (int)$row;
        $sheet->setCellValue($cellId, $column->Header);
        $sheet->getColumnDimensionByColumn($index)->setAutoSize(true);
        $style = $sheet->getStyle($cellId);
        $style->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $style->getFont()->setBold(true);
        $style->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB('EFEFEF');

    }
}

/**
 * @param PHPExcel_Worksheet $sheet
 * @param array $records
 * @param callable[] $formatters Key-value array with formatters
 * @param int $row First row in sheet for records
 * @return int
 */
function records($sheet, $records, $formatters = [], $row = 3)
{
    foreach ($records as $record) {
        $column = 0;
        $sheet->getRowDimension($row)->setRowHeight(24);

        foreach ($record as $key => $value) {
            $formattedValue = isset($formatters[$key]) ? $formatters[$key]($value) : $value;
            $style = $sheet->getStyleByColumnAndRow($column, $row);
            $style->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $style->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                ->getStartColor()->setRGB('FFFFFF');
            applyBorder($style);
            $sheet->setCellValueByColumnAndRow($column, $row, $formattedValue);

            $column++;
        }

        $row++;
    }

    return $row - 1;
}

function addEmptyRows($sheet, $row, $columnsCount)
{
    $sheet->mergeCellsByColumnAndRow(0, $row, $columnsCount - 1, $row);
    $style = $sheet->getStyleByColumnAndRow(0, $row);
    $sheet->getRowDimension($row)->setRowHeight(12);
    $sheet->setCellValueByColumnAndRow(0, $row, '');

    $style->getFill()->getStartColor()->setRGB('FFFFFF');
}

/**
 * @param PHPExcel_Worksheet $sheet
 * @param array $columns
 * @param int $row
 * @return void
 */
function footer($sheet, $columns, $row)
{
    $sheet->getRowDimension($row)->setRowHeight(24);

    // Title
    footerTitle($sheet, $row);

    // Values
    for ($columnId = 1; $columnId < count($columns); $columnId++) {
        footerValue($sheet, $columnId, $row, $columns[$columnId]->footer);
    }
}

/**
 * @param PHPExcel_Worksheet $sheet
 * @param int $row
 * @return void
 */
function footerTitle($sheet, $row)
{
    // Title
    $sheet->setCellValueByColumnAndRow(0, $row, 'TOTAIS');
    $style = $sheet->getStyleByColumnAndRow(0, $row);

    $style->getFont()->setBold(true);
    $style->getAlignment()
        ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)
        ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
}

function applyBorder($style)
{
    $styleArray = [
        'borders' => [
            'top' => [
                'style' => PHPExcel_Style_Border::BORDER_THICK,
                'color' => ['rgb' => 'EFEFEF']
            ],
            'bottom' => [
                'style' => PHPExcel_Style_Border::BORDER_THICK,
                'color' => ['rgb' => 'EFEFEF']
            ]
        ]
    ];

    $style->applyFromArray($styleArray);
}

/**
 * @param PHPExcel_Worksheet $sheet
 * @param $columnId
 * @param int $row
 * @param mixed $value
 * @return void
 */
function footerValue($sheet, $columnId, $row, $value)
{
    $sheet->setCellValueByColumnAndRow($columnId, $row, moedafooter($value));
    $style = $sheet->getStyleByColumnAndRow($columnId, $row);

    $style->getFont()->setBold(true);
    $style->getAlignment()
        ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)
        ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

}

/**
 * @param PHPExcel_Worksheet $sheet
 * @param string $loja
 */
function detalhesContrato($sheet, $loja)
{
    $json = getDetalhesContrato($loja);
    title($sheet, count($json->summary->columnsName), 'Detalhes do Contrato');
    columns($sheet, $json->summary->columnsName);
    $records = records($sheet, $json->records);
    addEmptyRows($sheet, ++$records, count($json->summary->columnsName));
    addEmptyRows($sheet, ++$records, count($json->summary->columnsName));
    return $records+1;
}

/**
 * @param PHPExcel_Worksheet $sheet
 * @param string $loja
 * @return void
 */
function aluguelContratual($sheet, $loja, $rows)
{
    $json = getAluguelContratual($loja);
    title($sheet, count($json->summary->columnsName), 'Aluguel Contratual', $rows++);
    columns($sheet, $json->summary->columnsName, $rows++);
    $records = records(
        $sheet,
        $json->records,
        [
            'Aluguel Contratual' => [Formatter::class, 'moeda'],
        ],
        $rows
    );

    addEmptyRows($sheet, ++$records, count($json->summary->columnsName));
    addEmptyRows($sheet, ++$records, count($json->summary->columnsName));
    return $records + 1;
}

/**
 * @param PHPExcel_Worksheet $sheet
 * @param string $loja
 * @return void
 */
function percentualVenda($sheet, $loja, $rows)
{
    $json = getPercentualVenda($loja);

    title($sheet, count($json->summary->columnsName), 'Aluguel %', $rows++);
    columns($sheet, $json->summary->columnsName, $rows++);
    $records = records(
        $sheet,
        $json->records,
        [
            'Volume Venda' => [Formatter::class, 'moeda'],
            '% Venda' => [Formatter::class, 'moeda'],
        ],
        $rows++
    );

    addEmptyRows($sheet, ++$records, count($json->summary->columnsName));
    addEmptyRows($sheet, ++$records, count($json->summary->columnsName));
    return $records + 1;
}

/**
 * @param PHPExcel_Worksheet $sheet
 * @param string $loja
 * @return void
 */
function reducaoAluguel($sheet, $loja, $rows)
{
    $json = getReducaoAluguel($loja);
    title($sheet, count($json->summary->columnsName), 'Redução de Aluguel Mínimo', $rows++);
    columns($sheet, $json->summary->columnsName, $rows++);
    return records(
        $sheet,
        $json->records,
        [
            'Valor' => [Formatter::class, 'moeda'],
            // $tbody .= '<td style="text-align:center; border-bottom: 1px solid #EFEFEF; height: 40px; vertical-align: middle;" data-element-type="long-text"> ' . $v['Observação'] . ' </td>';
        ],
        $rows++
    );
}

/**
 * @param PHPExcel_Worksheet $sheet
 * @param string $loja
 * @param string $mes
 * @param string $kpi
 * @return void
 */
function valorMensal($sheet, $loja, $mes, $kpi, $rows)
{
    $json = getValorMensal($loja, $mes, $kpi);
    title($sheet, count($json->summary->columnsName), 'Tabela Mensal', $rows++);
    columns($sheet, $json->summary->columnsName, $rows++);

    $formatters = [];
    foreach ($json->summary->columnsName as $column) {
        $formatters[$column->Header] = [Formatter::class, 'moedaIfNumeric'];
    }

    $currentRow = $rows;

    foreach ($json->records as $shopRecords) {
        $lastRow = records($sheet, $shopRecords, $formatters, $currentRow);
        $currentRow = $lastRow + 1;
    }

    footer($sheet, $json->summary->columnsName, $currentRow);

    addEmptyRows($sheet, ++$currentRow, count($json->summary->columnsName));
    addEmptyRows($sheet, ++$currentRow, count($json->summary->columnsName));

    return $currentRow+1;
}

/**
 * @param PHPExcel_Worksheet $sheet
 * @param string $loja
 * @param string $mes
 * @param string $kpi
 * @return void
 */
function visaoMeses($sheet, $loja, $mes, $kpi, $rows)
{
    $json = getVisaoMeses($loja, $mes, $kpi);
    title($sheet, count($json->summary->columnsName), 'Visão 12 Meses', $rows++);
    columns($sheet, $json->summary->columnsName, $rows++);

    $formatters = [];
    foreach ($json->summary->columnsName as $column) {
        $formatters[$column->Header] = [Formatter::class, 'moedaIfNumeric'];
    }

    // API returns array with one element
    $records = isset($json->records[0]) ? [$json->records[0]] : [];
    $currentRow = $rows;

    foreach ($records as $shopRecords) {
        $lastRow = records($sheet, $shopRecords, $formatters, $currentRow);
        $currentRow = $lastRow + 1;
    }
    footer($sheet, $json->summary->columnsName, $currentRow);

    addEmptyRows($sheet, ++$currentRow, count($json->summary->columnsName));
    addEmptyRows($sheet, ++$currentRow, count($json->summary->columnsName));

    return $currentRow + 1;
}

/**
 * @param PHPExcel_Worksheet $sheet
 * @param string $loja
 * @param string $mes
 * @param string $kpi
 * @return void
 */
function shoppings($sheet, $loja, $mes, $kpi, $rows)
{
    $json = getShoppings($loja, $mes, $kpi);
    $currentRow = $rows;

    $formatters = [];
    foreach ($json->summary->columnsName as $column) {
        $formatters[$column->Header] = [Formatter::class, 'verificaNuloIfNumeric'];
    }

    // Intermediate object contains on row, where key is shop name and value is shop records
    foreach ($json->records as $intermediateObject) {
        foreach ($intermediateObject as $shopName => $shopRecords) {
            title($sheet, count($json->summary->columnsName), $shopName, $currentRow);
            $currentRow++;
            columns($sheet, $json->summary->columnsName, $currentRow);
            $currentRow++;
            $lastRow = records($sheet, $shopRecords, $formatters, $currentRow);
            $currentRow = $lastRow + 1;

            // Footer for each shop
            $sheet->getRowDimension($currentRow)->setRowHeight(24);
            footerTitle($sheet, $currentRow);

            $shopFooterValues = searchShopFooter($json->recordsShopFooter, $shopName);
            $counter = 0; // Counter needed bacuse shopFooterValues is object
            foreach ($shopFooterValues as $value) {
                if ($counter < 2) {
                    // Skip first two fields
                    $counter++;
                    continue;
                }

                // 6 - is a columns shift
                footerValue($sheet, $counter + 6, $currentRow, $value);
                $counter++;
            }

            $currentRow++;

            addEmptyRows($sheet, ++$currentRow, count($json->summary->columnsName));
            addEmptyRows($sheet, ++$currentRow, count($json->summary->columnsName));
        }
    }
}

/**
 * @param array $records
 * @param string $shopName
 * @return object|array
 */
function searchShopFooter($records, $shopName)
{
    foreach ($records as $intermediateObject) {
        foreach ($intermediateObject as $recordsShopName => $recordsValues) {
            if ($recordsShopName === $shopName) {
                return isset($recordsValues[0]) ? $recordsValues[0] : [];
            }
        }
    }

    return [];
}
