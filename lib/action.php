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
            if(move_uploaded_file($uploaded_file,$this->filePath.$fileName.'.xlsx')) {
                $this->saveData($fileName);
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
    //保存excel数据到json数据
    public function saveData($fileName){
        $excelData = $this->getExcelData($fileName);

    }

    public function getExcelData($fileName){
        //$objReader = PHPExcel_IOFactory::createReaderForFile($this->filePath.$fileName.'.xlsx');
        $objPHPExcel = PHPExcel_IOFactory::load($this->filePath.$fileName.'.xlsx');
        //$objPHPExcel = $objReader->load($this->filePath.$fileName.'.xlsx');
        $objPHPExcel->setActiveSheetIndex(0);
        //$data = $objPHPExcel->getActiveSheet()->getCell('E3')->getValue();

     $objWorksheet = $objPHPExcel->getActiveSheet();
     $i = 0;
     echo "<table>";
     foreach($objWorksheet->getRowIterator() as $row){

         echo "<tr>";

             $cellIterator = $row->getCellIterator();
             $cellIterator->setIterateOnlyExistingCells(false);

             if( $i == 0 ){
                 echo '<thead>';
             }
             foreach($cellIterator as $cell){

                 echo '<td width="100%">' . $cell->getValue() . '</td>';

             }
             if( $i == 0 ){
                 echo '</thead>';
             }
             $i++;

         echo "</tr>";

     }
        echo "</table>";
        echo "<pre>";
        //print_r($data);
        return true;
    }


}