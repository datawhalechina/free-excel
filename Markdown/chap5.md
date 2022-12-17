# 第五章 Excel函数-文本函数

## 1.Text函数

Text函数可以将数值转换为指定格式的文本，其语法格式为TEXT(value,format_text)

【TEXT函数】=TEXT(值，自定义数字格式代码)

### 案例1

打开`data/chap5/5.1.xlsx`，点击【案例1】，将客户的消费日期和消费金额转成大写

针对遇到的问题，那么在D2单元格中输入

**=TEXT(C2,"[DBNUM2]")**            ----> **注意，这里的逗号要使用英文的逗号**

在E2单元格中输入

**=TEXT(A2,"[DBNUM1]yyyy年m月d日")**

DBNUM1和DBNUM2为2种常见的中文格式，一般金额用DBNUM2，日期用DBNUM1

![5.1](.\src\chap5\5.1.gif)

### 案例2

打开`data/chap5/5.1.xlsx`，点击【案例2】，将客户的消费日期转换为周次，即星期几

针对遇到的问题，那么在D2单元格中输入

**=TEXT(A2,"aaaa")**          

![5.2](.\src\chap5\5.2.gif)

### 案例3

打开`data/chap5/5.1.xlsx`，点击【案例3】，取客户消费的年、月、日

针对遇到的问题，那么在D2单元格中输入

**=TEXT(A2,"yyyy")**     或者 **=TEXT(A2,"e")** 

在E2单元格中输入

**=TEXT(A2,"m")**  或者 **=TEXT(A2,"mm")**         注意.这2种格式是有区别的

在F2单元格中输入

**=TEXT(A2,"d")**  或者 **=TEXT(A2,"dd")** 

![5.3](.\src\chap5\5.3.gif)

## 2.mid函数

打开`data/chap5/5.2.xlsx`，点击【案例1】，提取身份证中的生日，并转换成2022年12月4日这种格式

面对这个问题，可以使用mid函数提取生日

【MID函数】=MID(text,start_num,num_chars)

test:为要提取的文本字符串

start_num:为文本中要提取的第一个字符串的位置

num_chars为提取字符串的长度

因此可以在B2中输入 因为生日是8位数字，所以最后一个参数填8

**=MID(A2,7,8)**

在C2中输入格式化的生日

**=TEXT(MID(A2,7,8),"0000年00月00日")**

问题：这里TEXT函数格式为什么没有使用yyyy年mm月dd日 这种格式？

因为MID(A2,7,8)提取出来的是文本，不是日期，因此Excel无法识别日期的年月，所以用数字格式进行代替

![5.4](.\src\chap5\5.4.gif)

文本提取中相似的函数有LEFT，RIGHT

LEFT函数，以字符串左侧为起始位置，返回指定数量的字符

【LEFT函数】=MID(text,,num_chars)

text:要提取的字符串或单元格引用；
num_chars:要提取的字符数量

RIGHT函数，从字符串右侧首字符开始，从右向左提取指定的字符，其功能和LEFT函数完全一样，只是方向不同

【RIGHT函数】=MID(text,num_chars)

text:要提取的字符串或单元格引用；
num_chars:要提取的字符数量

## 3.replace函数

打开`data/chap5/5.2.xlsx`，点击【案例2】，现在需要将Excel表格打印，为了不泄露客户电话号码，需要将电话后5位进行屏蔽

REPLACEI函数作用：把一个文本字符串，人为指定一个位置，用定个数新字符进行替换。

【REPLACEI函数】=REPLACEI(old_text,start_num,num_chars,new_text)

old_text:需要替换的文本

start_num:需要替换文本的开始位置

num_chars:替换文本的长度

new_text:替换内容

因此可以在B2单元格中输入

**=REPLACE(A2,11,5,"#####")**

![5.5](.\src\chap5\5.5.gif)

## 练习

1.完成`data/chap5/5.1.xlsx`工作簿中的【案例1-3】

2.完成`data/chap5/5.2.xlsx`工作簿中的【案例1】中身份证后6位加密



