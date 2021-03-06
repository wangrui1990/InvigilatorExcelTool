## 背景
需求参考：excel数据自动化处理求助（用VBA或者Python或者其他都可以）
https://www.52pojie.cn/thread-1460112-1-1.html

## 技术
基于.Net Core WPF开发。

## 第一版本
初版已经完成了，界面如下：
![4.png](images/4.png)
导出的表格如下，样式也按照要求进行了调整：
![5.png](images/5.png)

![6.png](images/6.png)

## 需要注意
有以下几个需要注意的地方，为了开发方便，需要修改原始Excel中的一下几个地方。
1.        同一表格，不能有重名的列，故考试种类修改为“预设考试种类”
![1.png](images/1.png)
2.        处理外语较为麻烦，故需将外语拆分为日语和英语
![2.png](images/2.png)
3.        同一表格，不能有重名的列，故，监考老师需按如下修改列名：
![3.png](images/3.png)

## 其他
1.        程序做的匆忙，并未进行异常处理。需严格按照上诉要求进行修改，否则出错会直接崩溃。一旦出错，需重新启动程序。
2.        程序较为清晰，而且很多地方都预留了参数可配置，单元格样式配置，监考老师的分配规则，教室的分配规则等，可根据需要灵活修改。

## 操作流程
![3.gif](images/3.gif)

## 程序下载：
链接: https://pan.baidu.com/s/1H4SEXcUbs9Bq24x64gv4Ug 提取码: fczj[/md]

## 源码
https://github.com/wangrui1990/InvigilatorExcelTool
