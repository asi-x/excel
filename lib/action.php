 <?php

/**
 * Created by PhpStorm.
 * User: XFour
 * Date: 17/11/3
 * Time: 下午3:43
 */
class action
{
    private $fileExtend = 'xlsx';

    private $excelObj = null;
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
                mkdir($this->filePath,0777);
            }
            //文件后缀名
            $extension = substr($_FILES['file']['name'], strrpos($_FILES['file']['name'], '.')+1);
            $this->fileExtend = $extension;
            if(move_uploaded_file($uploaded_file,$this->filePath.$fileName.'.'.$extension)) {
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
    public function updateValue(){
        $sourceData = $this->excelTodata('Source.'.$this->fileExtend);
        //var_dump($sourceData);
        $targetData = $this->excelTodata('Target.'.$this->fileExtend);
        //var_dump($targetData);
        foreach ($sourceData as $row=>$rowItem){
            foreach ($rowItem  as $clumn=>$value){
                $v = trim($value);
                if(!empty($v) && preg_match('/^[0-9]+(.[0-9]{1,2})?$/',trim($v))){
                    $val = str_replace(',', '', $targetData[$row][$clumn])+$v;
                    $this->setValue($clumn.$row,$val,'Target.'.$this->fileExtend);
                }
            }
        }
        $this->getExcelObj('Target.'.$this->fileExtend)->getActiveSheet()->setTitle();
        $this->getExcelObj('Target.'.$this->fileExtend)->setActiveSheetIndex(0);
        $objWriter = PHPExcel_IOFactory::createWriter($this->getExcelObj('Target.'.$this->fileExtend), 'Excel2007');
        $objWriter->save($this->filePath.'Result.xlsx');//文件保存路径
        Header("Location: {$this->filePath}Result.xlsx");
        return true;

    }
    //保存excel数据读成数组,生成新的execl
    public function saveData($fileName,$extension){
        $excelData = $this->getExcelData($fileName,$extension);
        $this->arrToExcel($excelData,'',true);
    }

    private function excelTodata($fileName){
        $PHPExcel = $this->getExcelObj($fileName);

        $sheetData = $PHPExcel->getActiveSheet(0)->toArray(null,true,true,true);

        return $sheetData;
    }
    private function setValue($cell,$value,$filename){

        $PHPExcel = $this->getExcelObj($filename);
        return $PHPExcel->getActiveSheet(0)->setCellValue($cell, $value);
    }
    private function getExcelObj($filename){
        if(isset($this->excelObj[$filename]) && !empty($this->excelObj[$filename])){
            return $this->excelObj[$filename];
        }
        if($this->fileExtend == 'xls'){
            $objReader = PHPExcel_IOFactory::createReader('Excel5');
            $this->excelObj[$filename] = $objReader->load($this->filePath.$filename); // 文档名称
        }elseif($this->fileExtend == 'xlsx'){
            $objReader = PHPExcel_IOFactory::createReader('Excel2007');
            $this->excelObj[$filename] = $objReader->load($this->filePath.$filename); // 文档名称
        }else{
            return new stdClass();
        }
        return $this->excelObj[$filename];
    }





}