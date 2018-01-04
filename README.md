## javaExcel  

read and write Excel file by Java, includes .xls and .xlsx; Java Excel export;  
Java Excel processing based on Apache POI  
    
    
java 读写 `Excel` 文件,包括 `.xls` 和 `.xlsx` 文件; Java `Excel` 文件导出; java 基于  
`Apache POI ` 的 `Excel` 文件处理工具  
`Apache POI document: ` [https://poi.apache.org/spreadsheet/quick-guide.html](https://poi.apache.org/spreadsheet/quick-guide.html "https://poi.apache.org/spreadsheet/quick-guide.html")  
    
    
## 依赖(Denpency)  
	<!--poi is for Microsoft Excel 97 (-2003),poi-ooxml is for Microsoft Excel XML (2007+)-->
    <dependency>
      <groupId>org.apache.poi</groupId>
      <artifactId>poi</artifactId>
      <version>3.17</version>
    </dependency>
    <dependency>
      <groupId>org.apache.poi</groupId>
      <artifactId>poi-ooxml</artifactId>
      <version>3.17</version>
    </dependency>  
    
    
## 功能(Function)  

### read  

- read local Excel file, and return result by list  
- the program has progressed the blank cell and blank row of the Excel file  
- the result of iterating the Excel is `List<String[][]>`,and the `String[][]` is  
    a relular Two-dimensional array,it maybe has blank lines.  
    Demo: `com.ljq.utils.ExcelUtilTest`  
    
    
### write  

- create local Excel file and write data to it  
- the exported Excel file path must be an the only valid ,it can't be null,or a  
    directory,or an exist file,and it must be ended with ".xls" or ".xlsx"  
    
    
### 读取  

- 读取本地 `Excel`文件,并通过 `List` 集合返回
- 程序中包含对空白单元格、空白行的处理  
- 返回结果为:`List<String[][]>`,`String[][]` 是一个规则的二维数组,其中可能包含空白行,  
    关于返回结果的处理可以参考 test部分(`com.ljq.utils.ExcelUtilTest`)  
    
    
### 写入/导出  

- 创建本地 `Excel` 文件,并将数据写入  
- 所指定创建的 `Excel` 文件的路径必须是唯一的有效的,不能为空,不能是目录(文件夹),不能是已经  
    存在的文件,必须以 `.xls` 或者 `.xlsx` 结尾  
    
    

 

