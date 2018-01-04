package com.ljq.utils;

import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

public class ExcelUtilTest {

    @Test
    public void writeExcelFile1() throws Exception {

        // create cells
        String[][] strArr1 = {{"aaa","bbb","ccc"},{"ddd","eee","fff"},{"ggg","hhh","iii"}};
        String[][] strArr2 = {{"aa","bb","cc"},{"dd","ee","ff"}};
        List<String[][]> cellList = new ArrayList<String[][]>();
        cellList.add(strArr1);
        cellList.add(strArr2);

        // create sheet names
        List<String> sheetNameList = new ArrayList<String>();
//        sheetNameList.add("SheetDemo1");
//        sheetNameList.add("SheetDemo2");

        // exported Excel file path
//        String outExcelPath = "src\\resources\\out\\outExcelDemo-1.xls";
        String outExcelPath = "src\\resources\\out\\outExcelDemo-2.xlsx";

        // write data to Excel file
        ExcelUtil.writeExcelFile(cellList,sheetNameList,outExcelPath);


    }


    @Test
    public void writeExcelFile() throws Exception {

        String outExcelPath = "src\\resources\\out\\outExcel-1.xlsaaaa";
        ExcelUtil.writeExcelFile(null,null,outExcelPath);


    }

    @Test
    public void readExcelFile() throws Exception {

        String excelPath1 = "src\\resources\\excel\\demo-1.xlsx";
        String excelPath2 = "src\\resources\\excel\\demo-2.xls";

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
                System.out.println("------ cut-off line ------");
            }
        }



    }

}