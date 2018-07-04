package com.excel.demo.ExcelTest;

import org.apache.poi.ss.usermodel.DateUtil;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author guochunyuan
 * @create on  2018-06-07 14:03
 */
public class Main {
    public static void main(String[] args) throws Exception {
       String filePath = "test.xlsx";
        File file = new File(filePath);
        OutputStream out = new FileOutputStream(file);
        List<String> list = new ArrayList<>();
        String[] headers = {"Content-Disposition", "attachment; filename=ssss.xlsx"};
        String[] clouns = {};
        ExcelUtils.expoortExcelx("表头",headers,clouns,list,out,"yyyy-MM-dd HH:mm:ss");
    }
}
