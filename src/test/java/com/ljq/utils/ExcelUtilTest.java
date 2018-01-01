package com.ljq.utils;

import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

public class ExcelUtilTest {
    @Test
    public void readExcelFile() throws Exception {

        String excelPath1 = "D:\\develop\\repository\\git\\javaExcel\\src\\resources\\excel\\demo-1.xlsx";
        String excelPath2 = "D:\\develop\\repository\\git\\javaExcel\\src\\resources\\excel\\demo-2.xls";

        List<String[][]> list = new ArrayList<String[][]>();
        list = ExcelUtil.readExcelFile(excelPath1);

        // the result of iterating over the local excel file
        if(list != null && !list.isEmpty()){
            for (int i = 0; i < list.size(); i++) {
                String[][] strArr = list.get(i);
                for (int j = 0; j < strArr.length; j++) {
                    for (int k = 0; k < strArr[0].length; k++) {
                        // ignore the null row
                        if(strArr[j] != null){
                            System.out.print(strArr[j][k] + "\t");
                        }
                    }
                    System.out.println();
                }
            }
        }



    }

}