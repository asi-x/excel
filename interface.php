<?php
/**
 * Created by PhpStorm.
 * User: XFour
 * Date: 17/11/3
 * Time: 下午3:36
 */
error_reporting(E_ALL);
ini_set('display_errors', '1');
include "./lib/action.php";
$post = $_POST?$_POST:[];
$get = $_GET?$_GET:[];
$req = array_merge($post,$get);
if(!empty($req)){
    $action = new action();
    switch ($req['command']){
        case 'uploadSource':
            $data =  $action->uploadSource();
            break;
        case 'getExcelData':
            $data =  $action->getExcelData('Source');
            break;
        case 'uploadTarget':
            $data =  $action->uploadTarget();
            break;
        case 'updateValue':
            $data =  $action->updateValue();
            break;
        default:
            $result['data'] = '无效的command';
            break;
    }


}else{
    $result['data'] = '参数不能为空';
}
