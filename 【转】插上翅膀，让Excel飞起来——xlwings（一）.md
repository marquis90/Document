# 【转】插上翅膀，让Excel飞起来——xlwings（一）
python操作Excel的模块，网上提到的模块大致有：xlwings、xlrd、xlwt、openpyxl、pyxll等，他们提供的功能归纳起来有两种：一、用python读写Excel文件，实际上就是读写有格式的文本文件，操作excel文件和操作text、csv文件没有区别，Excel文件只是用来储存数据。二、除了操作数据，还可以调整Excel文件的表格宽度、字体颜色等。另外需要提到的是用COM调用Excel的API操作Excel文档也是可行的，相当麻烦基本和VBA没有区别。

## xlwings的特色
* xlwings能够非常方便的读写Excel文件中的数据，并且能够进行单元格格式的修改
* 可以和matplotlib以及pandas无缝连接
* 可以调用Excel文件中VBA写好的程序，也可以让VBA调用用Python写的程序。
* 开源免费，一直在更新

## 基本操作
### 1、打开已保存的Excel文档
```
# 导入xlwings模块，打开Excel程序，默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
 import xlwings as xw
 app=xw.App(visible=True,add_book=False)
 app.display_alerts=False
 app.screen_updating=False
 # 文件位置：filepath，打开test文档，然后保存，关闭，结束程序
 filepath=r'g:\Python Scripts\test.xlsx'
 wb=app.books.open(filepath)
 wb.save()
 wb.close()
 app.quit()
 ```
### 2、新建Excel文档，命名为test.xlsx，并保存在D盘。
 ```
 import xlwings as xw
 app=xw.App(visible=True,add_book=False)
 wb=app.books.add()
 wb.save(r'd:\test.xlsx')
 wb.close()
 app.quit()
 ```
### 3、在单元格输入值
新建test.xlsx，在sheet1的第一个单元格输入 “人生” ，然后保存关闭，退出Excel程序。
 ```
 import xlwings as xw
 app=xw.App(visible=True,add_book=False)
 wb=app.books.add()
 # wb就是新建的工作簿(workbook)，下面则对wb的sheet1的A1单元格赋值
 wb.sheets['sheet1'].range('A1').value='人生'
 wb.save(r'd:\test.xlsx')
 wb.close()
 app.quit()
 ```
### 4、打开已保存的test.xlsx，在sheet2的第二个单元格输入“苦短”，然后保存关闭，退出Excel程序
 ```
 import xlwings as xw
 app=xw.App(visible=True,add_book=False)
 wb=app.books.open(r'd:\test.xlsx')
 # wb就是新建的工作簿(workbook)，下面则对wb的sheet1的A1单元格赋值
 wb.sheets['sheet1'].range('A1').value='苦短'
 wb.save()
 wb.close()
 app.quit()
 ```
掌握以上代码，已经完全可以把Excel当作一个txt文本进行数据储存了，也可以读取Excel文件的数据，进行计算后，并将结果保存在Excel中。
