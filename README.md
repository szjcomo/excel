# szjcomo/excel/Excel类的使用说明
这是一个操作php操作excel的库，需要安装扩展依赖
## 安装 composer require szjcomo/excel
	构造函数
	construct($config = array())
|  参数名称 | 类型  |是否必传  | 默认值  |  说明 |
| ------------ | ------------ | ------------ | ------------ |
|  $conf | array  | 是  | array  |导出成excel文件是必传,导入excel不传|
|  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;path | string  | 否| ./static/excelport  | excel文件路径,需要导出  |