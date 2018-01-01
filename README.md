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

- 读取本地 `Excel`文件,并通过 `List` 集合返回 (read local Excel file, and return result  
    by List)  
- 程序中包含对空白单元格、空白行的处理 (the program has progressing the blank cell and  
    blank row of the Excel file)  
- 返回结果为:`List<String[][]>`,`String[][]` 是一个规则的二维数组,其中可能包含空白行,  
    关于返回结果的处理可以参考 test部分(`com.ljq.utils.ExcelUtilTest`) (the result  
    of iterating the Excel is `List<String[][]>`,and the `String[][]` is a relular  
    Two-dimensional array,it maybe has blank lines)  

 

