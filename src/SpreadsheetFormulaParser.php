<?php

namespace zkyg\SpreadsheetFormulaLaravel;


use Illuminate\Support\Str;
use PhpOffice\PhpSpreadsheet\Spreadsheet;


class SpreadsheetFormulaParser
{
    private $formula;
    private $values;
    private $excelFormula;
    private $translatedValues;
    private $generateFileDebug;


    function __construct(){
        $this->generateFileDebug = false;
    }

    /**
     * Shorthand for initialize and get class instance.
     * @return SpreadsheetFormulaParser
     */
    public static function getInstance(): SpreadsheetFormulaParser
    {
        return new SpreadsheetFormulaParser();
    }

    /**
     * Get result from supplied formula.
     * @param string $formula   Excel based formula
     * @param array $values     Array of values, index count have to equals with parameter in formula
     * @param bool $parseFormulaOnly    Debug only to return parsed formula and check result only
     * @param bool $outputFile  will write file to "storage/app/excel-formula" directory for debug purposes.
     * @return mixed|object
     * @throws \Exception
     */
    function calculate(string $formula, array $values, bool $parseFormulaOnly = false, bool $outputFile = false){
        $this->formula = $formula;
        $this->values = $values;
        $this->parseFormula();

        if ( $parseFormulaOnly ){
            return (object)[
                'originalFormula' => $this->formula,
                'excelFormula' => $this->excelFormula,
                'params' => $this->translatedValues
            ];
        }

        $this->generateFileDebug = $outputFile;
        return $this->excelFormulaSimulation();
    }

    /**
     * Parsing excel formula & validate params
     * @throws \Exception
     */
    private function parseFormula(){
        $match = [];
        preg_match_all('/(\[\[)([a-zA-Z])(\w+)(\]\])/',$this->formula, $match);
        $match = array_unique($match[0]); // complete tags in first(0) index then make em unique


        // make sure count of values and $match is equals
        if ( count($match) !== count($this->values) )
            throw new \Exception("Missmatch number of parameter. Formula expected " . count($match) . " but supplied : " . count($this->values), 400);

        $this->translatedValues = [];
        $count = 0;
        $this->excelFormula = $this->formula;
        foreach ( $match as $item ){
            $baseNameParam = str_replace("[[", '', str_replace("]]", '', $item));
            if ( !isset($this->values[$baseNameParam]) )
                throw new \Exception("Parameter with index '$baseNameParam' cannot be found.", 400);

            $cellAddr = self::getNameFromNumber($count) . "1";
            array_push($this->translatedValues, [
                'excelCoordinates' => $cellAddr,
                'value' => $this->values[$baseNameParam],
                'indexName' => $baseNameParam,
                'originalTag' => $item
            ]);

            $this->excelFormula = str_replace(" ", "", str_replace($item, $cellAddr, $this->excelFormula));
            $count++;
        }
    }

    /**
     * Do formula calculation and get the result
     * @return mixed
     * @throws \PhpOffice\PhpSpreadsheet\Calculation\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function excelFormulaSimulation(){

        if ( $this->generateFileDebug ){
            $temporaryPath = storage_path('app/excel-formula/');
            if ( !is_dir($temporaryPath) )
                @mkdir($temporaryPath);

            // check once again
            if ( !is_dir($temporaryPath) )
                throw new \Exception("Cannot create temporary path for working file @ $temporaryPath", 500);
        }

        $spreadSheet = new Spreadsheet();
        $sheet = $spreadSheet->getActiveSheet();
        foreach ($this->translatedValues as $value){
            $sheet->setCellValue($value['excelCoordinates'], $value['value']);
        }
        // set the formula on row 2
        $sheet->setCellValue("A2", $this->excelFormula);

        // get calculated value
        $value = $sheet->getCell('A2')->getCalculatedValue();

        if ( $this->generateFileDebug ){
            $filePath = self::generateFilePath($temporaryPath, 'xlsx');
            $writer = new Xlsx($spreadSheet);
            $writer->save($filePath);
        }

        return $value;
    }

    /**
     * Helper to generate random file path(with checking)
     * @param string $path
     * @param string $extension
     * @return string
     */
    public static function generateFilePath(string $path, string $extension){
        $filePath =  $path . Str::random() . "." . $extension;
        if ( file_exists($filePath) )
            return self::generatedFilePath($path, $extension);

        return $filePath;
    }

    /**
     * Helper function for translating number to letter based on spreadsheed(ex: 0 => A, 1 => B, ... )
     * @param $num
     * @return string
     */
    public static function getNameFromNumber($num) {
        $numeric = $num % 26;
        $letter = chr(65 + $numeric);
        $num2 = intval($num / 26);
        if ($num2 > 0) {
            return getNameFromNumber($num2 - 1) . $letter;
        } else {
            return $letter;
        }
    }
}
