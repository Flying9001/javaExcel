package com.ljq.utils;

import org.junit.Test;

import static org.junit.Assert.*;

public class FileUtilTest {

    @Test
    public void checkFilePath() throws Exception {

//        String filePath = "";
//        String filePath = "src\\resour";
//        String filePath = "src\\resources\\out\\outExcel-1.xlsaaaaa";
//        String filePath = "src\\resources\\out";
//        String filePath = "src\\resources\\out\\";
        String filePath = "src\\resources\\out\\outExcel-1.xls";


        System.out.println(FileUtil.checkFilePath(filePath));

    }

}