<?php
/**
 * |-----------------------------------------------------------------------------------
 * @Copyright (c) 2014-2018, http://www.sizhijie.com. All Rights Reserved.
 * @Website: www.sizhijie.com
 * @Version: 思智捷管理系统 1.5.0
 * @Author : como 
 * 版权申明：szjshop网上管理系统不是一个自由软件，是思智捷科技官方推出的商业源码，严禁在未经许可的情况下
 * 拷贝、复制、传播、使用szjshop网店管理系统的任意代码，如有违反，请立即删除，否则您将面临承担相应
 * 法律责任的风险。如果需要取得官方授权，请联系官方http://www.sizhijie.com
 * |-----------------------------------------------------------------------------------
 */
require '../vendor/autoload.php';
use szjcomo\excel\Excel;


//数据导入示例代码
$obj = new Excel();
$result = $obj->import('123.xls');
print_r($result);


//数据导出示例代码
/**
 * 注意事项:
 * xlswriter扩展只支持导出成功xlsx后缀的excel
 */
$obj1 = new Excel(['path'=>'./']);
$callbackClass = new ExcelCallback();
$field = [
	'id'=>['name'=>'序号'],
	'uname'=>['name'=>'学生姓名','callback'=>[$callbackClass,'unameHandler']],
	'add_time'=>['name'=>'报名时间','callback'=>[$callbackClass,'addTimeHandler']],
];
$data = [
	['id'=>1,'uname'=>'szjcomo','add_time'=>time()]
];
$result = $obj1->field($field)->data($data)->export('456.xlsx');
print_r($result);
