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
* [xlwings官方文档](http://docs.xlwings.org/en/stable/quickstart.html)  
* [插上翅膀，让Excel飞起来——xlwings (二)](http://www.jianshu.com/p/b534e0d465f7)           
* [插上翅膀，让Excel飞起来——xlwings (三)](http://www.jianshu.com/p/de7efe591c12)  
* [插上翅膀，让Excel飞起来——xlwings (四)](http://www.jianshu.com/p/7d6f53e3e6e9)        
* [excel中想实现使用Python代替VBA，请问应该怎么做？](https://www.zhihu.com/question/37937045)    
* [python模块:win32com用法详解](https://www.2cto.com/kf/201206/137809.html)    
* [python中使用xlrd、xlwt操作excel表格详解](http://www.jb51.net/article/60510.htm)    


在上一篇插上翅膀，让Excel飞起来——xlwings（一）中提到利用xlwings模块，用python操作Excel有如下的优点：

* xlwings能够非常方便的读写Excel文件中的数据，并且能够进行单元格格式的修改
* 可以和matplotlib以及pandas无缝连接
* 可以调用Excel文件中VBA写好的程序，也可以让VBA调用用Python写的程序。
* 开源免费，一直在更新
*本文紧接着上文介绍了xlwings模块一些常用的api

## 常用函数和方法

### 1.Book 工作簿常用的api
wb=xw.books[‘工作簿名称']  
wb.activate()激活为当前工作簿
wb.fullname 返回工作簿的绝对路径
wb.name 返回工作簿的名称
wb.save(path=None) 保存工作簿，默认路径为工作簿原路径，若未保存则为脚本所在的路径  
-wb. close() 关闭工作簿
代码例子：
```
    # 引用Excel程序中，当前的工作簿
    wb=xw.books.acitve
    # 返回工作簿的绝对路径
    x=wb.fullname
    # 返回工作簿的名称
    x=wb.name
    # 保存工作簿，默认路径为工作簿原路径，若未保存则为脚本所在的路径
    x=wb.save(path=None)
    # 关闭工作簿
    x=wb.close()
```
### 2、sheet 常用的api
``` 
   # 引用某指定sheet
   sht=xw.books['工作簿名称'].sheets['sheet的名称']
   # 激活sheet为活动工作表
   sht.activate()
   # 清除sheet的内容和格式
   sht.clear()
   # 清除sheet的内容
   sht.contents()
   # 获取sheet的名称
   sht.name
   # 删除sheet
   sht.delete
```
### 3、range常用的api
```
   # 引用当前活动工作表的单元格
   rng=xw.Range('A1')
   # 加入超链接
   # rng.add_hyperlink(r'www.baidu.com','百度',‘提示：点击即链接到百度')
   # 取得当前range的地址
   rng.address
   rng.get_address()
   # 清除range的内容
   rng.clear_contents()
   # 清除格式和内容
   rng.clear()
   # 取得range的背景色,以元组形式返回RGB值
   rng.color
   # 设置range的颜色
   rng.color=(255,255,255)
   # 清除range的背景色
   rng.color=None
   # 获得range的第一列列标
   rng.column
   # 返回range中单元格的数据
   rng.count
   # 返回current_region
   rng.current_region
   # 返回ctrl + 方向
   rng.end('down')
   # 获取公式或者输入公式
   rng.formula='=SUM(B1:B5)'
   # 数组公式
   rng.formula_array
   # 获得单元格的绝对地址
   rng.get_address(row_absolute=True, column_absolute=True,include_sheetname=False, external=False)
   # 获得列宽
   rng.column_width
   # 返回range的总宽度
   rng.width
   # 获得range的超链接
   rng.hyperlink
   # 获得range中右下角最后一个单元格
   rng.last_cell
   # range平移
   rng.offset(row_offset=0,column_offset=0)
   #range进行resize改变range的大小
   rng.resize(row_size=None,column_size=None)
   # range的第一行行标
   rng.row
   # 行的高度，所有行一样高返回行高，不一样返回None
   rng.row_height
   # 返回range的总高度
   rng.height
   # 返回range的行数和列数
   rng.shape
   # 返回range所在的sheet 
   rng.sheet
   #返回range的所有行
   rng.rows
   # range的第一行
   rng.rows[0]
   # range的总行数
   rng.rows.count
   # 返回range的所有列
   rng.columns
   # 返回range的第一列
   rng.columns[0]
   # 返回range的列数
   rng.columns.count
   # 所有range的大小自适应
   rng.autofit()
   # 所有列宽度自适应
   rng.columns.autofit()
   # 所有行宽度自适应
   rng.rows.autofit()
```
### 3、books 工作簿集合的api
```
   # 新建工作簿
   xw.books.add()
   # 引用当前活动工作簿
   xw.books.active
```
### 4、sheets 工作表的集合
```
   # 新建工作表
   xw.sheets.add(name=None,before=None,after=None)
   # 引用当前活动sheet
   xw.sheets.active
```
## 实例

大Z老师，教了小z同学怎么用python操作Excel之后，利用第一篇和第二篇的知识，编写了一个python小脚本，给小Z同学演示了一下怎么用python调整单元格的行宽、列宽和背景色，做一些Interesting的事。  

小Z同学在看了这么cliche但是好玩的东西之后，自己果断地修改了代码，改变了单元格的颜色，并在sheet里面进行了题字，然后，便有新的作品：  
下一课有机会教小z同学，利用python自带的time模块，让Excel中静态的画和字动起来，成为像gif一样的图片。
![ss](http://upload-images.jianshu.io/upload_images/2979196-a1a5011dd2410a59.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)
