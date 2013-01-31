<?php
//Snippet: showXlsx
//Modify by ashap
//pavel@pe-art.ru
//To use: create TV - name "xlsFile", type "File" 
//Only for .xslx file format!
//You must have a "PHP Excel library" in assets/libs catalog
//Params: &showBorder = `true|false` - show or not table border
require_once 'assets/libs/Classes/PHPExcel/IOFactory.php';
$showBorder = isset($showBorder) ? $showBorder : $showBorder;
switch ($showBorder)
{
case 'true' : $showBorder = '2';
break;
case 'false' : $showBorder = '1';
break;
}
//получить имя файла с таблицей
$xlsx = $modx->getObject('modTemplateVar', array('name' =>'xlsxFile'));
$xlsx = $xlsx->getValue($modx->resource->get('id'));
//имя файла с таблицей в формате html
$htm=str_replace('.xlsx','.html',$xlsx);
//читаем данные из xlsx файла
$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$objPHPExcel = $objReader->load($xlsx);
// и записываем их в html
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'HTML');
$objWriter->setSheetIndex(0);
$objWriter->save(str_replace('.xlsx', '.html', $xlsx));
// получаем данные из html файла
$table=file_get_contents($htm);
// получаем и обрабатываем таблицу стилей
$style=preg_match('|(<style [^>]+>[^<]+<\/style>)|',$table,$matches);
$style=str_replace(array('display:none;','visibility:hidden','*display:none','visibility:collapse;'),'',$matches[$showBorder]);

//получаем собственно cаму таблицу и выводим её
$output=substr($table,strpos($table,'<body>'),strpos($table,'</body>')-strpos($table,'<body>'));
$output=str_replace('<body>','',$output);
$output=str_replace('border="0" cellpadding="0" cellspacing="0"','border="1" cellpadding="1" cellspacing="1"',$output);
//$header=$modx->getObject('modChunk',array('name'=>'tableHeader'));
return $style.$output;
?>