package com.ljq.utils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

/**
 * @description: read and write excel files
 * @author: lujunqiang
 * @email: flying9001@gmail.com
 * @date: 2017/12/30
 */
public class ExcelUtil {

    // value "true" for DEBUG
    private static final boolean DBG = false;

    /**
     * iterator over local excel file(include .xls and .xlsx)
     * @param excelPath local excel file path
     *
     * @reruen list result of itetroing
     * */
    public static List<String[][]> readExcelFile(String excelPath){
        try {
            FileInputStream inputStream = new FileInputStream(excelPath);
            return readExcelFile(inputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return null;

    }

    /**
     * iterator over Excel file from stream
     * @param inputStream Excel file stream
     *
     * @return list result of itetroing
     * */
    public static List<String[][]> readExcelFile(FileInputStream inputStream){
        // result of iterating over sheets of excel file
        List<String[][]> list = new ArrayList<String[][]>();
        try {
            Workbook wb = WorkbookFactory.create(inputStream);
            // iterating over the excel file
            for (Sheet sheet : wb) {
                // Decide which rows to process
                int rowStart = sheet.getFirstRowNum();
                int rowEnd = sheet.getLastRowNum();
                if(rowEnd != 0){
                    // total number of row
                    int rowCount = rowEnd + 1;
                    if(DBG){System.out.println("rowCount: " + rowCount); }
                    // total number of column
                    int colCount = sheet.getRow(rowStart).getPhysicalNumberOfCells();
                    if(DBG){System.out.println("cosCount: " + colCount); }

                    // the result of one sheet(rows and cells)
                    String[][] strArr = new String[rowCount][colCount];

                    for (int rowNum = rowStart; rowNum < rowCount; rowNum++) {
                        Row row = sheet.getRow(rowNum);
                        if (row == null) {
                            if(DBG){System.out.println("this row is empty");}

                            strArr[rowNum] = null;
                            continue;
                        }else{
                            for (int colNum = 0; colNum < colCount; colNum++) {
                                Cell cell = row.getCell(colNum, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                                strArr[rowNum][colNum] = getCellValue(cell);
                            }
                        }
                    }
                    list.add(strArr);
                }
                if(DBG){System.out.println("----- cut-off line ------"); }
            }
            return list;
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        return null;

    }

    /**
     *  get String value of Excel cell
     *  @param cell Excel cell
     *
     *  @return string
     * */
    private static String getCellValue(Cell cell){
        if (cell == null) {
            if(DBG){System.out.println("this cell is empty");}
            return null;
        } else {
            String cellValue = "";
            switch (cell.getCellTypeEnum()) {
                case STRING:
                    cellValue = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        cellValue = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss:SSS").format(cell.getDateCellValue());
                    } else {
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                case BOOLEAN:
                    cellValue = String.valueOf(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    try {
                        cellValue = cell.getStringCellValue();
                    } catch (Exception e) {
                        try {
                            cellValue = String.valueOf(cell.getNumericCellValue());
                        } catch (Exception e1) {
                            cellValue = String.valueOf(cell.getCellFormula());
                        }
                    }
                    break;
                case BLANK:
                    cellValue = "";
                    break;
                case _NONE:
                    cellValue = "";
                    break;
                default:
                    cellValue = "";
                    break;
            }
            if(DBG){
                System.out.println("cellValue: " + cellValue);
            }
            return cellValue;
        }
    }



}
