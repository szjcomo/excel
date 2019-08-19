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
namespace szjcomo\excel;
/**
 * 本类用到了php组件
 * 请用 pecl install xlswriter命令进行安装 
 * 添加 extension = xlswriter.so 到 ini 配置
 * 文档地址：https://gitee.com/viest/php-ext-xlswriter#PECL
 */
Class Excel {
	/**
	 * [$header 单元格的头部信息]
	 * @var array
	 */
	Public $header = [];
	/**
	 * [$data 需要保存的数据]
	 * @var array
	 */
	Public $data = [];
	/**
	 * [$sheetName 表名称]
	 * @var string
	 */
	Public $sheetName = 'sheet1';
	/**
	 * [$config 数据导出保存路径]
	 * @var array
	 */
	Public $config = ['path'=>'./static/excelport'];
	/**
	 * [$excelObj excel操作对象]
	 * @var null
	 */
	Private $excelObj = null;
	/**
	 * [$fieldsCallback 设置字段回调函数]
	 * @var array
	 */
	Public $fieldsCallback = [];
	/**
	 * [__construct 构造函数]
	 * @Author    como
	 * @DateTime  2019-08-17
	 * @copyright 思智捷管理系统
	 * @version   [1.5.0]
	 */
	Public function __construct($config = null){
		if(!is_null($config) && is_array($config) && isset($config['path'])) $this->config = array_merge($this->config,$config);
		if(extension_loaded('xlswriter')){
			$this->excelObj = new \Vtiful\Kernel\Excel($this->config);
		}
	}
	/**
	 * [setSheetName 设置表名称]
	 * @Author    como
	 * @DateTime  2019-08-17
	 * @copyright 思智捷管理系统
	 * @version   [1.5.0]
	 * @param     [type]     $sheetName [description]
	 */
	Public function sheet($sheetName = null){
		if(!is_null($sheetName)) 
			$this->sheetName = $sheetName;
		return $this;
	}
	/**
	 * [fields 设置导出字段的回调函数]
	 * @Author    como
	 * @DateTime  2019-08-17
	 * @copyright 思智捷管理系统
	 * @version   [1.5.0]
	 * @param     array      $fields [description]
	 * @return    [type]             [description]
	 */
	Public function field($fields = []){
		if(!empty($fields)){
			$this->fieldsCallback = array_merge($this->fieldsCallback,$fields);
		}
		return $this;
	}
	/**
	 * [export 导出成excel文件,只支持xlsx格式 不支持xls格式]
	 * @Author    como
	 * @DateTime  2019-08-16
	 * @copyright 思智捷管理系统
	 * @version   [1.5.0]
	 * @param     array      $data      [description]
	 * @param     string     $saveName  [description]
	 * @param     string     $sheetName [description]
	 * @return    [type]                [description]
	 */
	Public function export($saveName = ''){
		if(empty($saveName)) $saveName = date('YmdHis').mt_rand(100000,999999).'.xlsx';
		$checkResult = $this->checkExportParams();
		if($checkResult['err'] !== false) return $checkResult;
		try{
			return $this->setHeaderHandler($this->fieldsCallback)->exportHandler($saveName);
		} catch(\Exception $err){
			return self::appResult($err->getMessage());
		}
	}
	/**
	 * [setHeaderHandler 处理头部信息]
	 * @Author    como
	 * @DateTime  2019-08-19
	 * @copyright 思智捷管理系统
	 * @version   [1.5.0]
	 */
	Protected function setHeaderHandler($data = []){
		if(!empty($data)){
			foreach($data as $key=>$val){
				if(!empty($val['name'])){
					$this->header[] = $val['name'];
				}
			}
		}
		return $this;
	}
	/**
	 * [exportHandler 实现真正的处理导出逻辑]
	 * @Author    como
	 * @DateTime  2019-08-17
	 * @copyright 思智捷管理系统
	 * @version   [1.5.0]
	 * @param     [type]     $saveName [description]
	 * @param     [type]     $callback [description]
	 * @return    [type]               [description]
	 */
	Protected function exportHandler($saveName){
		$fileObj = $this->excelObj->constMemory($saveName,$this->sheetName);
		try{
			$index = 0;
			if(!empty($this->header)) {
				$index = 1;
				$fileObj->header($this->header);
			}
			foreach($this->data as $key=>$val){
				$columnIndex = 0;
				foreach($this->fieldsCallback as $k=>$v){
					$value = isset($val[$k]) && !empty($val[$k])?$val[$k]:'';
					$this->callbackHandler($fileObj,$index,$columnIndex,$value,$k,$val);
					$columnIndex++;
				}
				$index++;
			}
			$action = $fileObj->output();
			return self::appResult('SUCCESS',$action,false);
		} catch(\Exception $err){
			return self::appResult($err->getMessage());
		}
	}
	/**
	 * [callbackHandler 导出时处理回调函数]
	 * @Author    como
	 * @DateTime  2019-08-17
	 * @copyright 思智捷管理系统
	 * @version   [1.5.0]
	 * @param     [type]     $fileObj     [description]
	 * @param     [type]     $index       [description]
	 * @param     [type]     $columnIndex [description]
	 * @param     [type]     $v           [description]
	 * @param     [type]     $k           [description]
	 * @param     [type]     $val         [description]
	 * @return    [type]                  [description]
	 */
	Protected function callbackHandler($fileObj,$index,$columnIndex,$v,$k,$val){
		if(isset($this->fieldsCallback[$k]) && is_array($this->fieldsCallback[$k]) && isset($this->fieldsCallback[$k]['callback'])){
			call_user_func($this->fieldsCallback[$k]['callback'],$fileObj,$index,$columnIndex,$v,$k,$val);
		} else {
			$fileObj->insertText($index,$columnIndex,$v);
		}			
	}
	/**
	 * [checkExportParams 检测导出参数]
	 * @Author    como
	 * @DateTime  2019-08-17
	 * @copyright 思智捷管理系统
	 * @version   [1.5.0]
	 * @return    [type]     [description]
	 */
	Public function checkExportParams(){
		if(empty($this->data)){
			return self::appResult('没有需要导出的数据,请设置需要导出的数据内容...');
		}
		if(empty($this->excelObj)){
			return self::appResult('您当前没有安装xlswriter扩展,请先安装扩展,具体详情请查看api');
		}
		if(empty($this->data[0]) || !is_array($this->data[0])){
			return self::appResult('数据格式不正确,数据格式必须是二维数组');
		}
		return self::appResult('SUCCESS',null,false);
	}
	/**
	 * [appResult 统一返回格式]
	 * @Author    como
	 * @DateTime  2019-08-17
	 * @copyright 思智捷管理系统
	 * @version   [1.5.0]
	 * @param     string     $info [description]
	 * @param     [type]     $data [description]
	 * @param     boolean    $err  [description]
	 * @return    [type]           [description]
	 */
	Public static function appResult($info = '',$data = null,$err = true){
		return ['info'=>$info,'data'=>$data,'err'=>$err];
	}
	/**
	 * [getExcelObj description]
	 * @Author    como
	 * @DateTime  2019-08-17
	 * @copyright 思智捷管理系统
	 * @version   [1.5.0]
	 * @return    [type]     [description]
	 */
	Public function data($data = []){
		$this->data = $data;
		return $this;
	}


	/**
	 * [import 导入工作表]
	 * @Author    como
	 * @DateTime  2019-08-19
	 * @copyright 思智捷管理系统
	 * @version   [1.5.0]
	 * @param     string     $fileName [description]
	 * @param     boolean    $is_data  [description]
	 * @param     [type]     $debug    [description]
	 * @return    [type]               [description]
	 */
	Public static function import($fileName = '',$index = 0,$is_data = false,$debug = null){
		$obj = null;
		if($is_data) {
			$obj = \SimpleXLSX::parse($fileName,true,$debug);
		} else {
			$checkResult = self::importHandler($fileName);
			if($checkResult['err'] !== false) return $checkResult;
			switch ($checkResult['data']) {
				case 'xls':
					$obj = \SimpleXLS::parse($fileName,$is_data,$debug);
					break;
				default:
					$obj = \SimpleXLSX::parse($fileName,$is_data,$debug);
					break;
			}
		}
		if($obj->success()){
			$data = $obj->rows($index);
			return self::appResult('SUCCESS',$data,false);
		} else {
			$info = 'ERROR';
			switch($checkResult['data']){
				case 'xls':
					$info = $obj->error();
					break;
				default:
					$info = $obj->parseError();
			}
			return self::appResult($info,null);
		}
	}

	/**
	 * [importHandler 导入工作表时的检测工作]
	 * @Author    como
	 * @DateTime  2019-08-19
	 * @copyright 思智捷管理系统
	 * @version   [1.5.0]
	 * @param     string     $fileName [description]
	 * @return    [type]               [description]
	 */
	Protected static function importHandler($fileName = ''){
		if(empty($fileName)) return self::appResult('需要导入的excel工作薄名称不能为空');
		if(!file_exists($fileName)) return self::appResult('需要导入的excel工作薄不存在,请检查...');
		$data = pathinfo($fileName);
		$importType = ['xls','xlsx'];
		if(!in_array($data['extension'], $importType)) return self::appResult('需要导入的文件类型不合法');
		return self::appResult('SUCCESS',$data['extension'],false);
	}
}