package com.gzmut.office.util;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.Map;

/**
 * excel工具类
 * @author MXDC
 * @date 2019/7/10
 **/
public class ExcelUtils {
    /** 文件名称*/
    private static  String FILENAME="";
    /** XSSFWorkbook工作簿对象*/
    private static XSSFWorkbook workBook;

    /** 放回一個工作簿對象*/
    public static XSSFWorkbook getWorkBook() {
        return workBook;
    }

    /**
     * 設置一個工作簿
     * @param fileName 工作簿的絕對地址
     */
    public static void setWorkook(String fileName) {
        FILENAME = fileName;
        // 判断字符串是否为空
        if (FILENAME==null || "".equals(FILENAME)){
            return ;
        }
        // 文件流
        InputStream stream = null;
        try {
            stream = new FileInputStream(new File(FILENAME));
            // excel为2003，进行转换
            if (FILENAME.endsWith(".xls")){
                // 进行转换2003，转2007
            } else if (FILENAME.endsWith(".xlsx")){
                workBook = new XSSFWorkbook(stream);
            } else {
                // LOGGER.debug("此文件{}不是word文件", path);
                // 这里要logger
                System.out.printf("此文件不是word文件");
            }
        }catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (null != stream){
                try {
                    stream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                    // LOGGER.debug("读取excel文件失败");
                    System.out.printf("读取excel文件失败");
                }
            }
        }
        return ;
    }
    /**
     * 判断是表名是否存在
     * @param sheetName
     * @return
     */
    public static boolean checkSheetName(String sheetName){
        // 判断workbook是否为空
        if (workBook == null){
            return false;
        }
        // 通过获取表判断是否存在表名
        return  workBook.getSheet(sheetName) == null ? false : true;
    }

    public static Map<Integer, Map<Integer,String>> getRegionCellValue(String sheetName,CellRangeAddress cellRangeAddress){
        // 判断workbook,sheetName,cellRangeAddress是否为空
        if (workBook == null || sheetName == null || "".equals(sheetName) || cellRangeAddress == null){
            System.out.println("工作簿，表名或地址域为空");
            return null;
        }
        //获取表
        XSSFSheet sheet = workBook.getSheet(sheetName);
        if ( sheet == null){
            System.out.println("表格为空");
            return null;
        }
        // cellMap用于存放某个单元格或某个区域的值
        Map<Integer,Map<Integer,String>> cellMap = new HashMap<Integer,Map<Integer,String>>();
        // 遍历地址，取出单元格的值存放在map中
        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();
        for(int i = firstColumn; i <= lastRow; i++){
            Map colMap = new HashMap<Integer,String>();
            for (int j = firstColumn; j <lastColumn ; j++) {
                colMap.put(j,getCellValueAsString(getCell(sheet,new CellAddress(i,j))));
            }
            cellMap.put(i,colMap);
        }
        return cellMap;
    }

    /**
     * 获取单元格各类型值，返回字符串类型
     * @param cell excel 单元格
     * @return java.lang.String
     */
    public static String getCellValueAsString(Cell cell){
        // 判断是否为null或空串
        if (cell==null || cell.toString().trim().equals("")) {
            return "";
        }
        String cellValue = "";
        // 获取单元格类型名称
        CellType cellType = cell.getCellType();
        switch (cellType){
            case STRING : //字符
                cellValue= cell.getStringCellValue().trim();
                cellValue= StringUtils.isEmpty(cellValue) ? "" : cellValue;
                break;
            case BOOLEAN: //布尔值
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case NUMERIC: //数值,日期,货币等也为数值
                if (HSSFDateUtil.isCellDateFormatted(cell)) {  //判断日期类型
                    cellValue =    String.valueOf(cell.getDateCellValue());
                } else {  //否
                    cellValue = new DecimalFormat("#.######").format(cell.getNumericCellValue());
                }
                break;
            case ERROR: // 异常
                break;
            case _NONE: // 未知类型
                break;
            case FORMULA: // 公式
                break;
            case BLANK: // 空白
                break;
            default:
                cellValue = "";
                break;
        }
        return cellValue;
    }

    /**
     * 根据一个地址返回一个单元格
     * @param sheet 一个excel表格
     * @param cellAddress 一个单元格地址
     * @return org.apache.poi.ss.usermodel.Cell
     */
    public static Cell getCell(XSSFSheet sheet, CellAddress cellAddress){
        if (sheet == null || cellAddress == null){
            return null;
        }
        return  sheet.getRow(cellAddress.getRow()).getCell(cellAddress.getColumn());
    }
}
