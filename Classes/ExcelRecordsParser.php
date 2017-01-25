<?php
declare(strict_types = 1);

namespace Smichaelsen\ExcelRecordsParser;

class ExcelRecordsParser
{

    public function parse(string $filepath, array $parsingConfiguration): \Generator
    {
        $excelObject = \PHPExcel_IOFactory::load($filepath);
        foreach ($excelObject->getAllSheets() as $worksheet) {
            $i = 0;
            foreach ($worksheet->getRowIterator() as $row) {
                if (!empty($parsingConfiguration['skip_lines'])) {
                    if ($i++ < $parsingConfiguration['skip_lines']) {
                        continue;
                    }
                }
                $record = [];
                foreach ($parsingConfiguration['fields'] as $fieldName => $fieldParsingConfiguration) {
                    $possibleColumnNames = array_map('trim', explode('//', $fieldParsingConfiguration['column']));
                    $value = null;
                    $possibleColumnName = null;
                    foreach ($possibleColumnNames as $possibleColumnName) {
                        $cellName = $possibleColumnName . $row->getRowIndex();
                        $value = $worksheet->getCell($cellName)->getValue();
                        if (!empty($value)) {
                            break;
                        }
                    }
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
