<?php
require("PHPExcel.php");
//数据源的文件名
$fileName = "recruit.txt";
//生成的文件名
$purposeFileName = "众测用户招募表";
$purposeSheetNameT1 = "T1";
$purposeSheetNameT2 = "T2";
$purposeSheetNameU1 = "U1";
//各型号手机的IMEI的匹配规则
$t1ImeiMatchRule1 = "86451602.*";
$t1ImeiMatchRule2 = "86459302.*";
$t2ImeiMatchRule1 = "99000620.*";
$t2ImeiMatchRule2 = "99000621.*";
$u1ImeiMatchRule1 = "86579002.*";
$u1ImeiMatchRule2 = "86784002.*";
$u1ImeiMatchRule3 = "86784102.*";
$u1ImeiMatchRule4 = "99000716.*";
//初始化不同型号手机的容器
$arrayT1 = [];
$arrayT2 = [];
$arrayU1 = [];
$error = [];

//把数据根据手机型号分类
$myFile = fopen($fileName, "r") or die("Unable to open file!");
while(!feof($myFile)){
    $serializeString = fgets($myFile);
    $unserializeString = unserialize($serializeString);
    if(preg_match($t1ImeiMatchRule1,$unserializeString['imei']) || preg_match($t1ImeiMatchRule2,$unserializeString['imei'])){
        $arrayT1 = $unserializeString;
    }elseif(preg_match($t2ImeiMatchRule1,$unserializeString['imei']) || preg_match($t2ImeiMatchRule2,$unserializeString['imei'])){
        $arrayT2 = $unserializeString;
    }elseif(preg_match($u1ImeiMatchRule1,$unserializeString['imei']) || preg_match($u1ImeiMatchRule2,$unserializeString['imei']) || preg_match($u1ImeiMatchRule3,$unserializeString['imei']) || preg_match($u1ImeiMatchRule4,$unserializeString['imei'])){
        $arrayU1 = $unserializeString;
    }else{
        $error = $unserializeString;
    }
}

//根据各种手机型号的数据来创建电子表格
$objPHPExcel =  new PHPExcel();
fclose($myFile);
var_dump($error);
?>