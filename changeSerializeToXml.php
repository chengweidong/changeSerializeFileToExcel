<?php
require("PHPExcel.php");
//sheetIndex的索引
$sheetIndex = 0;
//数据源的文件名
$fileName = "recruit.txt";
//生成的文件名
$purposeFileName = "众测用户招募表.xls";
$purposeSheetNameT1 = "T1";
$purposeSheetNameT2 = "T2";
$purposeSheetNameU1 = "U1";
//各型号手机的IMEI的匹配规则
$t1ImeiMatchRule1 = "/^86451602.*/";
$t1ImeiMatchRule2 = "/^86459302.*/";
$t2ImeiMatchRule1 = "/^99000620.*/";
$t2ImeiMatchRule2 = "/^99000621.*/";
$u1ImeiMatchRule1 = "/^86579002.*/";
$u1ImeiMatchRule2 = "/^86784002.*/";
$u1ImeiMatchRule3 = "/^86784102.*/";
$u1ImeiMatchRule4 = "/^99000716.*/";
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
        $arrayT1[] = $unserializeString;
    }elseif(preg_match($t2ImeiMatchRule1,$unserializeString['imei']) || preg_match($t2ImeiMatchRule2,$unserializeString['imei'])){
        $arrayT2[] = $unserializeString;
    }elseif(preg_match($u1ImeiMatchRule1,$unserializeString['imei']) || preg_match($u1ImeiMatchRule2,$unserializeString['imei']) || preg_match($u1ImeiMatchRule3,$unserializeString['imei']) || preg_match($u1ImeiMatchRule4,$unserializeString['imei'])){
        $arrayU1[] = $unserializeString;
    }else{
        $error[] = $unserializeString;
    }
}

//根据各种手机型号的数据来创建电子表格
$objPHPExcel =  new PHPExcel();
//创建第一个表单
if(!empty($arrayT1)){
    $myWorkSheetT1 = new PHPExcel_Worksheet($objPHPExcel, 'T1众测用户信息');
    $objPHPExcel->addSheet($myWorkSheetT1, $sheetIndex);
    // $objPHPExcel->getSheet($sheetIndex);  //获得第一个表单
    $objPHPExcel->setActiveSheetIndex($sheetIndex++);  //获得第一个表单

    // $objPHPExcel->getActiveSheet()->setTitle('T1众测用户信息');

    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'bbsId');
    $objPHPExcel->getActiveSheet()->setCellValue('B1', 'cloudId');
    $objPHPExcel->getActiveSheet()->setCellValue('C1', 'imei');
    $objPHPExcel->getActiveSheet()->setCellValue('D1', 'email');
    $objPHPExcel->getActiveSheet()->setCellValue('E1', 'time');
    $objPHPExcel->getActiveSheet()->fromArray(
        $arrayT1,    // The data to set
        NULL,        // Array values with this value will not be set
        'A2'         // Top left coordinate of the worksheet range where
                     //    we want to set these values (default is A1)
    );
}
//创建第二个表单
if(!empty($arrayT2)){
    $myWorkSheetT2 = new PHPExcel_Worksheet($objPHPExcel, 'T2众测用户信息');
    $objPHPExcel->addSheet($myWorkSheetT2, $sheetIndex);
    // $objPHPExcel->getSheet($sheetIndex);  //获得第二个表单
    $objPHPExcel->setActiveSheetIndex($sheetIndex++);  //获得第二个表单

    // $objPHPExcel->createSheet();
    // $objPHPExcel->getSheet(++$sheetIndex);
    // $objPHPExcel->getActiveSheet()->setTitle('T2众测用户信息');

    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'bbsId');
    $objPHPExcel->getActiveSheet()->setCellValue('B1', 'cloudId');
    $objPHPExcel->getActiveSheet()->setCellValue('C1', 'imei');
    $objPHPExcel->getActiveSheet()->setCellValue('D1', 'email');
    $objPHPExcel->getActiveSheet()->setCellValue('E1', 'time');
    $objPHPExcel->getActiveSheet()->fromArray(
        $arrayT2,    // The data to set
        NULL,        // Array values with this value will not be set
        'A2'         // Top left coordinate of the worksheet range where
                     //    we want to set these values (default is A1)
    );
}
//创建第三个表单
if(!empty($arrayU1)){
    $myWorkSheetU1 = new PHPExcel_Worksheet($objPHPExcel, 'U1众测用户信息');
    $objPHPExcel->addSheet($myWorkSheetU1,$sheetIndex);
    // $objPHPExcel->getSheet($sheetIndex++);  //获得第三个表单
    $objPHPExcel->setActiveSheetIndex($sheetIndex++);  //获得第三个表单
    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'bbsId');
    $objPHPExcel->getActiveSheet()->setCellValue('B1', 'cloudId');
    $objPHPExcel->getActiveSheet()->setCellValue('C1', 'imei');
    $objPHPExcel->getActiveSheet()->setCellValue('D1', 'email');
    $objPHPExcel->getActiveSheet()->setCellValue('E1', 'time');
    $objPHPExcel->getActiveSheet()->fromArray(
        $arrayU1,    // The data to set
        NULL,        // Array values with this value will not be set
        'A2'         // Top left coordinate of the worksheet range where
                     //    we want to set these values (default is A1)
    );
}
//创建第四个表单
if(!empty($error)){
    $myWorkSheetError = new PHPExcel_Worksheet($objPHPExcel, '发生错误的众测用户信息');
    $objPHPExcel->addSheet($myWorkSheetError,$sheetIndex);
    // $objPHPExcel->getSheet($sheetIndex++);  //获得第三个表单
    $objPHPExcel->setActiveSheetIndex($sheetIndex++);  //获得第三个表单
    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'bbsId');
    $objPHPExcel->getActiveSheet()->setCellValue('B1', 'cloudId');
    $objPHPExcel->getActiveSheet()->setCellValue('C1', 'imei');
    $objPHPExcel->getActiveSheet()->setCellValue('D1', 'email');
    $objPHPExcel->getActiveSheet()->setCellValue('E1', 'time');
    $objPHPExcel->getActiveSheet()->fromArray(
        $error,      // The data to set
        NULL,        // Array values with this value will not be set
        'A2'         // Top left coordinate of the worksheet range where
                     //    we want to set these values (default is A1)
        );
}
//创建电子表格文件
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, "Excel5");
$objWriter->save($purposeFileName);

$objPHPExcel->disconnectWorksheets();
unset($objPHPExcel);
fclose($myFile);
?>