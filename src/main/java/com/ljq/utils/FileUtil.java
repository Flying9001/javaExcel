package com.ljq.utils;

import java.io.File;

/**
 * @description: file util class
 * @author: lujunqiang
 * @email: flying9001@gmail.com
 * @date: 2018/1/3
 */
public class FileUtil {

    /**
     * check out file path and return the type of File
     * @param filePath path of file
     *
     * return stirng type of File,like "file","directory","notFileOrDir"
     * */
    public static String checkFilePath(String filePath){

        if(filePath != null && !filePath.equals("")){
            File file = new File(filePath);
            if(file.isDirectory()){
                return "directory";
            }else{
                if(file.isFile()){
                    return "file";
                }
                return "notFileOrDir";
            }
        }
        return null;
    }


}
