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

    /** 返回一個工作簿對象*/
    public static XSSFWorkbook getWorkBook() {
        return workBook;
    }

    /**
     * 設置一個当前操作的工作簿
     * @param fileName 工作簿的絕對地址
     */
    public static void setWorkbook(String fileName) {
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
     * 判断表名是否存在
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

    /**
     * 获取一个单元格或区域的值
     * @param sheetName 表名
     * @param cellRangeAddress 一个单元格或访问 eg A1 | A2:C3
     * @return
     */
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
            for (int j = firstColumn; j <= lastColumn ; j++) {
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
        if (cell==null || "".equals(cell.toString().trim())) {
            return "";
        }
        String cellValue = "";
        // 获取单元格类型名称
        CellType cellType = cell.getCellType();
        switch (cellType){
            case STRING :
                cellValue= cell.getStringCellValue().trim();
                cellValue= StringUtils.isEmpty(cellValue) ? "" : cellValue;
                break;
            case BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            // 数值,日期,货币等也为数值
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    cellValue =    String.valueOf(cell.getDateCellValue());
                } else {
                    cellValue = new DecimalFormat("#.######").format(cell.getNumericCellValue());
                }
                break;
            case ERROR:
                break;
            // 未知类型
            case _NONE:
                break;
            // 公式
            case FORMULA:
                break;
            // 空白
            case BLANK:
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

    /**
     * 获取单元格字号
     * @param sheetname 表名
     * @param cell 单元格
     * @return 返回字号
     */
    public String getCellFontHeight(String sheetname,String cell){
        //  Short size   ;Get the font height in unit's of 1/20th of a point.
        return (new Short(workBook.getSheet(sheetname).
                getRow(getrow(cell)-1).getCell(getcolumn(cell)-1).
                getCellStyle().getFont().getFontHeight())).toString();
    }

    /**
     * 获取字体
     * @param sheetname 表名
     * @param cell 单元格
     * @return 返回字体
     */
    public String getCellFont(String sheetname,String cell){
        return workBook.getSheet(sheetname).getRow(getrow(cell)-1).getCell(getcolumn(cell)-1).getCellStyle().getFont().getFontName();
    }

    /**
     * 获取字体颜色
     * @param sheetname 表名
     * @param cell 单元格
     * @return 返回字体颜色
     */
    public String getCellFontColor(String sheetname,String cell){

        return rgbToString((workBook.getSheet(sheetname).getRow(getrow(cell)-1).getCell(getcolumn(cell)-1).getCellStyle().getFont().getXSSFColor().getRGB()));
    }






    /**
     *
     * @param cell 把cell（A3）转化为(1 3)
     * @return
     */
    public static int getcolumn(String cell){//cell="AC5"
        Map<Character,Integer> map = new HashMap();
        map.put('A',1);
        map.put('B',2);
        map.put('C',3);
        map.put('D',4);
        map.put('E',5);
        map.put('F',6);
        map.put('G',7);
        map.put('H',8);
        map.put('I',9);
        map.put('J',10);
        map.put('K',11);
        map.put('L',12);
        map.put('M',13);
        map.put('N',14);
        map.put('O',15);
        map.put('P',16);
        map.put('Q',17);
        map.put('R',18);
        map.put('S',19);
        map.put('T',20);
        map.put('U',21);
        map.put('V',22);
        map.put('W',23);
        map.put('X',24);
        map.put('Y',25);
        map.put('Z',26);
        int row = 0;
        String reg = "[^a-zA-Z]";
        cell = cell.replaceAll(reg,"");
        char[] cell1 = cell.toCharArray();
        for(int i = 0,j=cell1.length-1;i<cell1.length;i++,j--){
            row+=map.get(cell1[i])*Math.pow(26,j);
        }
        return row;
    }
    public static int getrow(String cell){
        String reg = "[!^a-zA-Z]";
        int num =(Integer.parseInt(cell.replaceAll(reg,"")));
        return num;
    }
    public String rgbToString(byte[] RGB){
        System.out.println(RGB);
        return new String(RGB);
    }
}
