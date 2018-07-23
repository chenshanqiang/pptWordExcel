<?php
//$srcfilename = 'F:/Apache/htdocs/pptWordExcel/test.xls';
//$destfilename = 'F:/Apache/htdocs/pptWordExcel/test.pdf';
//if (!file_exists($srcfilename)) {
//	return;
//}
//$excel = new \COM("excel.application") or die("Unable to instantiate excel");
//$workbook = $excel -> Workbooks -> Open($srcfilename, null, false, null, "1", "1", true);
//$workbook -> ExportAsFixedFormat(0, $destfilename);
//$workbook -> Close();
//$excel -> Quit();
$excelone = new \COM("excel.application") or die("Unable to instantiate excel");
echo $excelone;
?>
