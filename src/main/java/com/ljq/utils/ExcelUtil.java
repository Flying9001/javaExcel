package com.ljq.utils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
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
    private static final boolean DBG = true;

    /**
     * iterator over local excel file(include .xls and .xlsx)
     * @param excelPath local path of excel file
     *
     * @reruen list result of itetroing
     * */
    public static List<String[][]> readExcelFile(String excelPath){
        try {
            Workbook workbook = WorkbookFactory.create(new File(excelPath));
            return iteratorWorkBook(workbook);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * iterator over Excel file from stream(need more memory)
     * @param inputStream stream of Excel file
     *
     * @return list result of iterating
     * */
    public static List<String[][]> readExcelFile(FileInputStream inputStream){
        try {
            Workbook workbook = WorkbookFactory.create(inputStream);
            return iteratorWorkBook(workbook);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * write Excel file(.xls and xlsx)
     * checkout the path of exported Excel file
     * @param cellList  value list of Excel cells
     * @param sheetNameList names of Excel sheets
     * @param outExcelPath path of output Excel
     *
     * return boolean weather success writing Excel file to local
     *
     * */
    public static boolean writeExcelFile(List<String[][]> cellList,List<String> sheetNameList, String outExcelPath){
        Workbook workbook = null;
        if(outExcelPath != null && !outExcelPath.equals("")){
            File outExcelFile = new File(outExcelPath);
            if(!outExcelFile.isDirectory()){
                if(outExcelFile.isFile()){
                    if(DBG){ System.out.println("outExcelPath: " + outExcelPath); }

                    return true;
                }
                if(DBG){ System.out.println("outExcelPath: " + outExcelPath + " is not a file.."); }
                return false;
            }
            if(DBG){ System.out.println("outExcelPath: " + outExcelPath + " is a directory."); }
            return false;
        }
        return false;
    }












    /**
     * iterator over sheets from Excel file
     * @param workbook Excel file,contains .xls and .xlsx
     *
     * @return list result of iterating
     *
     * */
    private static List<String[][]> iteratorWorkBook(Workbook workbook){
        if(workbook == null){
            return null;
        }else{
            // result of iterating over sheets of excel file
            List<String[][]> list = new ArrayList<String[][]>();
            // iterating over the excel file
            for (Sheet sheet : workbook) {
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
        }



        return null;
    }

    /**
     *  get String value of Excel cell
     *  @param cell cell of Excel file
     *
     *  @return string value of cell
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
