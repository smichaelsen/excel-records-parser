<?php
declare(strict_types = 1);

namespace Smichaelsen\ExcelRecordsParser;

class ExcelRecordsParser
{

    public function parse(string $filepath, array $parsingConfiguration): \Generator
    {
        $excelObject = \PHPExcel_IOFactory::load($filepath);
        foreach ($excelObject->getAllSheets() as $worksheet) {
            foreach ($worksheet->getRowIterator() as $row) {
                $record = [];
                foreach ($parsingConfiguration as $fieldName => $fieldParsingConfiguration) {
                    $cellName = $fieldParsingConfiguration['column'] . $row->getRowIndex();
                    $value = $worksheet->getCell($cellName)->getValue();
                    if (is_callable($fieldParsingConfiguration['transform'])) {
                        $value = call_user_func($fieldParsingConfiguration['transform'], $value);
                    }
                    if ($value !== null) {
                        $record[$fieldName] = $value;
                    }
                }
                yield $record;
            }
        }
    }

}
