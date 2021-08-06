# spreadsheet-formula-excel

Library to parse and execute simple spreadsheet formula inside Laravel project.
Basically it's using PhpSpreadsheet to simulate and get the formula result.

- Install
```
composer require zkyg/spreadsheet-formula-laravel
```
  
- How to use
```
$formula = "=IF( [[value1]] < 5, \"ok\", [[value1]] + [[value2]] )";
$values = [
    'value1' => 7,
    'value2' => 3
];
$res = \zkyg\SpreadsheetFormulaLaravel\SpreadsheetFormulaParser::getInstance()->calculate(
    $formula,
    $values
);
// $res = 10
```
- Rules:
```
- Parameter count have to be equals in both formula & value list.
- Value index name is alphanumeric and cannot begin with number.
- Function segment separated with "," not ";"   
```
