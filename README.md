### Excel类使用说明

- 该组件主要是excel导入导出  只能导出为.xlsx格式的文件 不支持xls格式 导入时可导出xlsx xls

-------------

#### 命名空间：` szjcomo\excel\Excel ` 
                    
>excel 数据导出 依赖php扩展 xlswriter 如果您没有安装 请使用 pecl install xlswriter 后加入php.ini  添加 extension = xlswriter.so 到配置 文档地址：https://gitee.com/viest/php-ext-xlswriter#PECL

#### 方法列表：

|  类型 | 方法名称   | 参数说明  |  方法说明 |
| ------------ | ------------ | ------------ | ------------ |
| public  | construct()  | 请查看参数说明  | 构造函数  |
| public  | sheet()  | 请查看参数说明  | 设置sheet表名称  |
| public  | field()  | 请查看参数说明  | 导出字段设置 可使用回调函数  |
| public  | data()  | 请查看参数说明  | 需要导出的数据  |
| public  | export()  | 请查看参数说明  | 实现数据导出  |
| static  | import()  | 请查看参数说明  | 实现数据导入  |



#### 参数说明：
- 函数原型 construct($config = null)

|   参数名称| 参数类型  | 是否必传  |  备注 |
| ------------ | ------------ | ------------ | ------------ |
|  $config | array  | 是/否  | 导出时必传 |
|   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;path| string  | 是 | 导出时文件保存目录地址  |

- 函数原型 sheet($sheetName = null)

|   参数名称| 参数类型  | 是否必传  |  备注 |
| ------------ | ------------ | ------------ | ------------ |
|  $sheetName | string  | 是  | 表名称 |

- 函数原型 field($fields = [])

|   参数名称| 参数类型  | 是否必传  |  备注 |
| ------------ | ------------ | ------------ | ------------ |
|  $fields | array  | 是  | 字段数组 |
|  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;key | string  | 是  | 字段名称下标 |
|  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;value | array  | 是  | 字段值 |
|  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;name | string  | 是  | 字段下标显示名称 |

- 函数原型 data($data = [])

|   参数名称| 参数类型  | 是否必传  |  备注 |
| ------------ | ------------ | ------------ | ------------ |
|  $data | array  | 是  | 二维数组,设置需要导出的数据 |

- 函数原型 export($saveName = '')

|   参数名称| 参数类型  | 是否必传  |  备注 |
| ------------ | ------------ | ------------ | ------------ |
|  $saveName | string  | 是  | 需要保存的文件名 必须是.xlsx结尾 |

- 函数原型 import($fileName = '',$index = 0,$is_data = false,$debug = null)

|   参数名称| 参数类型  | 是否必传  |  备注 |
| ------------ | ------------ | ------------ | ------------ |
|  $fileName | string  | 是  | 需要导出入的文件路径 .xlsx .xls |
|  $index | int  | 否  | 默认导出第一个工作表 |
|  $is_data | bool  | 否  | $fileName是否是远程字符串 支持远程文件导入 默认false 本地文件 |
|  $debug | bool  | 否  | 调试错误信息 |



#### 使用示例：
######  一、数据导出
```php
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
```
######  二、数据导入
```php
$obj = new Excel();
$result = $obj->import('123.xls');
print_r($result);
```
