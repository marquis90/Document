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

## 引用工作簿、工作表和单元格

### 1、引用工作簿，注意工作簿应该首先被打开
 ```
    wb.=xw.books['工作簿的名字‘]
 ```
### 2、引用活动工作簿
```
    wb=xw.books.active
```
### 3、引用工作簿中的sheet
```
    sht=xw.books['工作簿的名字‘].sheets['sheet的名字']
    # 或者
    wb=xw.books['工作簿的名字']
    sht=wb.sheets[sheet的名字]
```
### 4、引用活动sheet
``` 
    sht=xw.sheets.active
```
### 5、引用A1单元格
``` 
    rng=xw.books['工作簿的名字‘].sheets['sheet的名字']
    # 或者
    sht=xw.books['工作簿的名字‘].sheets['sheet的名字']
    rng=sht.range('A1')
```
### 6、引用活动sheet上的单元格
```
    # 注意Range首字母大写
    rng=xw.Range('A1')
```
其中需要注意的是单元格的完全引用路径是：
```
    # 第一个Excel程序的第一个工作薄的第一张sheet的第一个单元格
    xw.apps[0].books[0].sheets[0].range('A1')
```
迅速引用单元格的方式是
```  
    sht=xw.books['名字'].sheets['名字']
    # A1单元格
    rng=sht[’A1']
    # A1:B5单元格
    rng=sht['A1:B5']
    # 在第i+1行，第j+1列的单元格
    # B1单元格
    rng=sht[0,1]
    # A1:J10
    rng=sht[:10,:10]
```
PS： 对于单元格也可以用表示行列的tuple进行引用
```
    # A1单元格的引用
    xw.Range(1,1)
    #A1:C3单元格的引用
    xw.Range((1,1),(3,3))
```
## 储存数据

### 1、储存单个值
```
    # 注意".value“
    sht.range('A1').value=1
```
### 2、储存列表
```
    # 将列表[1,2,3]储存在A1：C1中
    sht.range('A1').value=[1,2,3]
    # 将列表[1,2,3]储存在A1:A3中
    sht.range('A1').options(transpose=True).value=[1,2,3] 
    # 将2x2表格，即二维数组，储存在A1:B2中，如第一行1，2，第二行3，4
    sht.range('A1').options(expand='table')=[[1,2],[3,4]]
```
## 读取数据

### 1、读取单个值
```
    # 将A1的值，读取到a变量中
    a=sht.range('A1').value
```
### 2、将值读取到列表中
```
    #将A1到A2的值，读取到a列表中
    a=sht.range('A1:A2').value
    # 将第一行和第二行的数据按二维数组的方式读取
    a=sht.range('A1:B2').value
```
## 参考文献
*[xlwings官方文档](http://docs.xlwings.org/en/stable/quickstart.html)  
*[插上翅膀，让Excel飞起来——xlwings （二）](http://www.jianshu.com/p/b534e0d465f7)  
*[插上翅膀，让Excel飞起来——xlwings (三)](http://www.jianshu.com/p/de7efe591c12)  
*[插上翅膀，让Excel飞起来——xlwings (四)](http://www.jianshu.com/p/7d6f53e3e6e9)  
*[excel中想实现使用Python代替VBA，请问应该怎么做？](https://www.zhihu.com/question/37937045)  
*[python模块:win32com用法详解](https://www.2cto.com/kf/201206/137809.html)  
*[python中使用xlrd、xlwt操作excel表格详解](http://www.jb51.net/article/60510.htm)  
