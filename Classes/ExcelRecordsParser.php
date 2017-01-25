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
                    $value = $this->getValue($worksheet, $row, $fieldParsingConfiguration);
                    $value = $this->mapValue($value, $fieldParsingConfiguration);
                    $value = $this->transformValue($value, $fieldParsingConfiguration);
                    if ($value !== null) {
                        $record[$fieldName] = $value;
                    }
                }
                yield $record;
            }
        }
    }

    protected function getValue(\PHPExcel_Worksheet $worksheet, \PHPExcel_Worksheet_Row $row, array $fieldParsingConfiguration): string
    {
        $possibleColumnNames = array_map('trim', explode('//', $fieldParsingConfiguration['column']));
        $possibleColumnName = null;
        foreach ($possibleColumnNames as $possibleColumnName) {
            $cellName = $possibleColumnName . $row->getRowIndex();
            $value = $worksheet->getCell($cellName)->getValue();
            if (!empty($value)) {
                return $value;
            }
        }
        return '';
    }

    protected function mapValue(string $value, array $fieldParsingConfiguration): string
    {
        if (!isset($fieldParsingConfiguration['map'])) {
            return $value;
        }
        if ($fieldParsingConfiguration['map_mode'] === 'strict' && !in_array($value, array_keys($fieldParsingConfiguration['map']))) {
            throw new \Exception('Value "' . $value . '" could not be mapped. Strict mode."', 1485343840);
        }
        foreach ($fieldParsingConfiguration['map'] as $match => $replace) {
            if ($value === $match) {
                return $replace;
            }
        }
        return $value;
    }

    protected function transformValue(string $value, array $fieldParsingConfiguration): string
    {
        if (!is_callable($fieldParsingConfiguration['transform'])) {
            return $value;
        }
        return (string)call_user_func($fieldParsingConfiguration['transform'], $value);
    }

}
