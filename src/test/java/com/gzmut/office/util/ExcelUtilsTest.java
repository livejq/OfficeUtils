package com.gzmut.office.util;

import com.alibaba.fastjson.JSONObject;
import com.google.gson.JsonObject;
import netscape.javascript.JSObject;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
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
          // 设置一个工作簿
          ExcelUtils.setWorkook("E:\\Desktop\\计算机设备全年销量统计表.xlsx");
          // 获取一个工作簿
          XSSFWorkbook workBook = ExcelUtils.getWorkBook();
          // 获取一个表
          XSSFSheet sheet = workBook.getSheetAt(0);
          // 获取表里的一行
          XSSFRow row = sheet.getRow(1);
          // 获取一个单元格
          XSSFCell cell = row.getCell(0);
//        System.out.println(String.valueOf(cell.getDateCellValue()));
//        System.out.println(HSSFDateUtil.isCellDateFormatted(cell));
        // 获取单元格类型
        CellType cellType = cell.getCellType();
//        System.out.println(cellType);

        //int cellType=cell.getCellType();
        //将字符串解析为地址
        CellRangeAddress a1 = CellRangeAddress.valueOf("D4:F6");
//        a1.forEach(cellAddress -> sheet.getRow(cellAddress.getRow()).forEach(cell -> System.out.println(cell.getStringCellValue())));

        Map<Integer, Map<Integer, String>> regionCellValue = ExcelUtils.getRegionCellValue("销售情况", a1);
        System.out.println(regionCellValue.toString());
    }
}