

/**
 * 生成excel
 * @param array $data 数据源
 * @param array $col 列标题
 * @param array $row 行标题
 * @param boolean $show 是否输出（否为保存文件）
 */
 public function createExcel($data,$sheetname='sheet1',$filename='out',$col=array(),$row=array(),$show=true,$filedir=''){
 vendor('Excel.PHPExcel');
 vendor('Excel.PHPExcel.Writer.Excel5');
 $Excel=new PHPExcel();
 $ExcelWriter=new PHPExcel_Writer_Excel5($Excel);

$objActSheet = $Excel->getActiveSheet();
 $objActSheet->setTitle($sheetname);
 $colcount=count($col)-1;
 $rowcount=count($row);
 //设置列标题
 if($col){
 foreach($col as $k=>$v){
 $objActSheet->setCellValue(chr(65+$k).'1', $v);
 }
 }
 //设置行标题
 if($row){
 foreach($row as $k=>$v){
 $objActSheet->setCellValue(chr(65).($k+2),$v);
 }
 }
 $initrow=1;
 //填充内容
 foreach($data as $k=>$v){
 $initrow+=1;
 for($i=0;$i<$colcount;$i++){
 $val=isset($v[$i])?$v[$i]:'-';
 $objActSheet->setCellValue(chr(66+$i).($initrow),$val);
 }
 }
 $outputFileName = $filename.".xls";
 if($show){
 header("Content-Type: application/force-download");
 header("Content-Type: application/octet-stream");
 header("Content-Type: application/download");
 header('Content-Disposition:inline;filename="'.$outputFileName.'"');
 header("Content-Transfer-Encoding: binary");
 header("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT");
 header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
 header("Pragma: no-cache");
 $ExcelWriter->save('php://output');
 }
 else{
 $ExcelWriter->save($filedir.$outputFileName);
 }
 }


//调用方法：

$col[0]='日期\时间';
 for($i=1;$i<25;$i++){
 $col[]=$i-1;
 }
 $row=array_keys($data);
 $Excel=new ExcelAction();
 $Excel->createExcel($data,'registerdata','register'.time(),$col,$row);

