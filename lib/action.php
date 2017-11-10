 <?php

/**
 * Created by PhpStorm.
 * User: XFour
 * Date: 17/11/3
 * Time: 下午3:43
 */
class action
{
    public function __construct()
    {
        $this->filePath = './file/';
        require_once './lib/Classes/PHPExcel.php';
        require_once './lib/Classes/PHPExcel/Reader/Excel2007.php';
        require_once './lib/Classes/PHPExcel/Reader/Excel5.php';
        require_once './lib/Classes/PHPExcel/Writer/Excel2007.php';
        require_once './lib/Classes/PHPExcel/Writer/Excel5.php';
        require_once './lib/Classes/PHPExcel/IOFactory.php';
    }
    //上传
    public function upload($fileName){
        $result['msg'] = '';
        $result['status'] = false;

        //判断是否上传成功（是否使用post方式上传）
        if(is_uploaded_file($_FILES['file']['tmp_name'])) {
            //把文件转存到你希望的目录（不要使用copy函数）
            $uploaded_file=$_FILES['file']['tmp_name'];
            if(!file_exists($this->filePath)) {
                mkdir($this->filePath);
            }
            //文件后缀名
            $extension = substr($_FILES['file']['name'], strrpos($_FILES['file']['name'], '.')+1);
            if($extension != 'xls' && $extension != 'xlsx'){
                $result['msg'] = "文件格式错误";
                return $result;
            }
            if(move_uploaded_file($uploaded_file,$this->filePath.$fileName.'.'.$extension)) {
                $this->saveData($fileName,$extension);
                $result['status'] = true;
                $result['msg'] = $_FILES['file']['name']."上传成功";
                return $result;
            } else {
                $result['msg'] = "移动文件失败";
                return $result;
            }
        } else {
            $result['msg'] = "文件上传失败";
            return $result;
        }
    }

    //数据文件上传
    public function uploadSource(){
        return $this->upload('Source');
    }
    //结构文件上传
    public function uploadTarget(){
        return $this->upload('Target');
    }
    //保存excel数据读成数组,生成新的execl
    public function saveData($fileName,$extension){
        $excelData = $this->getExcelData($fileName,$extension);
        $this->exportExcel($excelData,'',true);
    }

    //excel文件生成数组
    public function getExcelData($fileName,$extension){

        if($extension == 'xls'){
            $objReader = PHPExcel_IOFactory::createReader('Excel5');
            $PHPExcel = $objReader->load($this->filePath.$fileName.'.xls'); // 文档名称
        }elseif($extension == 'xlsx'){
            $objReader = PHPExcel_IOFactory::createReader('Excel2007');
            $PHPExcel = $objReader->load($this->filePath.$fileName.'.xlsx'); // 文档名称
        }

        //工作表的数量
        $sheetCount = $PHPExcel->getSheetCount();
        //工作表名称数组
        $sheetNames = $PHPExcel->getSheetNames();
        $res = array();
        for ($SheetID = 0; $SheetID < $sheetCount; $SheetID++) {
            /**读取excel文件中的工作表*/
            $name = $sheetNames[$SheetID];
            $currentSheet = $PHPExcel->getSheetByName($name);
            $highestRow = $currentSheet->getHighestRow(); // 取得总行数
            $highestColumn = $currentSheet->getHighestColumn(); // 获得最后的列
            $highestColumnNum = PHPExcel_Cell::columnIndexFromString($highestColumn);//将列名转为数字
            for ($row = 1; $row <= $highestRow; $row++) {
                for ($column = 0; $column != $highestColumnNum; $column++) {
                    $val = $currentSheet->getCellByColumnAndRow($column, $row)->getValue();
                    if($val instanceof PHPExcel_RichText){
                        //富文本转换字符串
                        $val = $val->__toString();
                    }
                    $res[$name][$row-1][$column] = $val;
                }
            }

        }
        //print_r($res);die;
        return $res;

    }

    //数组生成excel文件
    function arrToExcel($list,$filename = '',$excel2007=false){

        if(empty($filename)) $filename = time();
        if(!is_array($list)) return false;

        //初始化PHPExcel()
        $objPHPExcel = new PHPExcel();
        //设置保存版本格式
        if($excel2007){
            $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
            $filename = $filename.'.xlsx';
        }else{
            $objWriter = new PHPExcel_Writer_Excel5($objPHPExcel);
            $filename = $filename.'.xls';
        }
        $startSheet = 0;
        foreach ($list as $key => $val) {
            //设置边框
            $styleArray = [
                'borders' => [
                    'allborders' => [
                        'style' => PHPExcel_Style_Border::BORDER_THIN,//细边框
                    ],
                ],
            ];
            $objPHPExcel->setActiveSheetIndex($startSheet); //切换到新创建的工作表
            $objActSheet = $objPHPExcel->getActiveSheet();
            $objActSheet->setTitle($key); //设置工作表名称
            if($startSheet < count($list)-1){
                $objPHPExcel->createSheet();//创建新的工作表
            }
            //数据写入1
            /*$startRow = 1;
            foreach ($val as $row) {
                foreach ($row as $key => $value){
                    $columnNum = PHPExcel_Cell::stringFromColumnIndex($key);
                    //这里是设置单元格的内容
                    $objActSheet->setCellValue($columnNum.$startRow,$value);
                    //设置背景颜色
                    $objActSheet->getStyle($columnNum.'1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $objActSheet->getStyle($columnNum.'1')->getFill()->getStartColor()->setARGB('B0C4DE');
                    //边框设置
                    $objActSheet->getStyle($columnNum.$startRow)->applyFromArray($styleArray);
                    //设置单元格为文本
                    //$objPHPExcel->getActiveSheet()->getStyle($header_arr[$key].$startRow)->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_TEXT);
                    //宽度
                    $objActSheet->getColumnDimension($columnNum)->setWidth(16);
                }
                $startRow++;
            }*/
            //数据写入2
            $objActSheet->fromArray($val, NULL, 'A1');
            $startRow = 1;
            foreach ($val as $row) {
                foreach ($row as $ke => $value) {
                    $header = PHPExcel_Cell::stringFromColumnIndex($ke);
                    //设置背景颜色
                    $objActSheet->getStyle($header.'1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $objActSheet->getStyle($header.'1')->getFill()->getStartColor()->setARGB('B0C4DE');
                    //边框设置
                    $objActSheet->getStyle($header.$startRow)->applyFromArray($styleArray);
                    //宽度
                    $objActSheet->getColumnDimension($header)->setWidth(16);
                }
                $startRow++;
            }
            $startSheet ++;
        }
        ob_end_clean();//清除缓冲区,避免乱码

        $objWriter->save('./file/'.$filename);

    }

    //生成excel并导出excel文件
    function exportExcel($list,$filename = '',$excel2007=false){

        if(empty($filename)) $filename = time();
        if(!is_array($list)) return false;

        //初始化PHPExcel()
        $objPHPExcel = new PHPExcel();
        //设置保存版本格式
        if($excel2007){
            $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
            $filename = $filename.'.xlsx';
        }else{
            $objWriter = new PHPExcel_Writer_Excel5($objPHPExcel);
            $filename = $filename.'.xls';
        }
        $startSheet = 0;
        foreach ($list as $key => $val) {
            //设置边框
            $styleArray = [
                'borders' => [
                    'allborders' => [
                        'style' => PHPExcel_Style_Border::BORDER_THIN,//细边框
                    ],
                ],
            ];
            $objPHPExcel->setActiveSheetIndex($startSheet); //切换到新创建的工作表
            $objActSheet = $objPHPExcel->getActiveSheet();
            $objActSheet->setTitle($key); //设置工作表名称
            if($startSheet < count($list)-1){
                $objPHPExcel->createSheet();//创建新的工作表
            }
            //数据写入
            $objActSheet->fromArray($val, NULL, 'A1');
            //单元格设置
            $startRow = 1;
            foreach ($val as $row) {
                foreach ($row as $ke => $value) {
                    $header = PHPExcel_Cell::stringFromColumnIndex($ke);
                    //设置背景颜色
                    $objActSheet->getStyle($header.'1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $objActSheet->getStyle($header.'1')->getFill()->getStartColor()->setARGB('B0C4DE');
                    //边框设置
                    $objActSheet->getStyle($header.$startRow)->applyFromArray($styleArray);
                    //宽度
                    $objActSheet->getColumnDimension($header)->setWidth(16);
                }
                $startRow++;
            }
            $startSheet ++;
        }

        // 下载这个表格，在浏览器输出
        ob_end_clean();//清除缓冲区,避免乱码
        header("Pragma: public");
        header("Expires: 0");
        header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
        header("Content-Type:application/force-download");
        header("Content-Type:application/vnd.ms-execl");
        header("Content-Type:application/octet-stream");
        header("Content-Type:application/download");;
        header('Content-Disposition:attachment;filename='.$filename.'');
        header("Content-Transfer-Encoding:binary");

        $objWriter->save('php://output');

    }




}