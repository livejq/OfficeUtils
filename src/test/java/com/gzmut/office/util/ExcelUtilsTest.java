package com.gzmut.office.util;

import com.alibaba.fastjson.JSONObject;
import com.google.gson.JsonObject;
import netscape.javascript.JSObject;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Set;

import static org.junit.Assert.*;

public class ExcelUtilsTest {
    @Test
    public void test() throws IOException {
//        FileInputStream fisSource = new FileInputStream(new File("E:\\Desktop\\计算机设备全年销量统计表.xlsx"));
//        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fisSource);
//        XSSFSheet sheet = xssfWorkbook.getSheet("平均价格");
//        XSSFWorkbook workbook = ExcelUtils.getWorkbook("E:\\Desktop\\计算机设备全年销量统计表.xlsx");
 //       System.out.println(ExcelUtils.checkSheetName(workbook, "销情况"));
            ExcelUtils.setWorkook("E:\\Desktop\\计算机设备全年销量统计表.xlsx");
        XSSFWorkbook workBook = ExcelUtils.getWorkBook();
        XSSFSheet sheet = workBook.getSheetAt(0);
        //将字符串解析为地址
        CellRangeAddress a1 = CellRangeAddress.valueOf("A2:A2");
        System.out.println(a1);
        System.out.println(a1.formatAsString());
        a1.forEach(cellAddress -> System.out.println(cellAddress.getColumn()));
    }
}