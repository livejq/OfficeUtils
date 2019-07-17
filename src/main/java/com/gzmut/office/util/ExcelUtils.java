package com.gzmut.office.util;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * excel工具类
 * @author MXDC
 * @date 2019/7/10
 **/
public class ExcelUtils {
    /** 文件名称*/
    private  String FILENAME="";
    /** XSSFWorkbook工作簿对象*/
    private XSSFWorkbook workBook;

    /** 放回一個工作簿對象*/
    public XSSFWorkbook getWorkBook() {
        return workBook;
    }

    /**
     * 設置一個工作簿
     * @param fileName 工作簿的絕對地址
     */
    public void setWorkbook(String fileName) {
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
                System.out.printf("此文件不是Excel文件");
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
    public boolean checkSheetName(String sheetName){
        // 判断workbook是否为空
        if (workBook == null){
            return false;
        }
        // 通过获取表判断是否存在表名
        return  workBook.getSheet(sheetName) == null ? false : true;
    }

    public Map<String, Map<String,String>> getCellValue(String sheetName,CellRangeAddress cellRangeAddress){
        // 判断workbook,sheetName,cellRangeAddress是否为空
        if (workBook == null || sheetName == null || "".equals(sheetName) || cellRangeAddress == null){
            return null;
        }
        //获取表
        XSSFSheet sheet = workBook.getSheet(sheetName);
        if ( sheet == null){
            return null;
        }

        return null;
    }

    /**
     *返回单元格内容
     * @param sheetname 表名
     * @param cell 单元格
     * @return 返回单元格内容
     */
    public String getCellValue (String sheetname,String cell){//A3
        // 判断workbook是否为空
        if (workBook == null){
            return "ERROR,WorkBook为空";
        }
        return workBook.getSheet(sheetname).getRow(getrow(cell)-1).getCell(getcolumn(cell)-1).getStringCellValue();
    }


    /**
     *返回字号
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
     *返回字体
     * @param sheetname 表名
     * @param cell 单元格
     * @return 返回字体
     */
    public String getCellFont(String sheetname,String cell){
        return workBook.getSheet(sheetname).getRow(getrow(cell)-1).getCell(getcolumn(cell)-1).getCellStyle().getFont().getFontName();
    }

    /**
     *返回字体颜色
     * @param sheetname 表名
     * @param cell 单元格
     * @return 返回字体颜色
     */
    public byte[] getCellFontColor(String sheetname,String cell){
        return workBook.getSheet(sheetname).getRow(getrow(cell)-1).getCell(getcolumn(cell)-1).getCellStyle().getFont().getXSSFColor().getRGB();
    }

    /**
     *返回单元格是否合并
     * @param sheetname 表名
     * @param cell 单元格
     * @return 返回单元格是否合并
     */
    public boolean isCellMerge(String sheetname,String cell){
        XSSFSheet sheet = workBook.getSheet(sheetname);
        int sheetMergeCount = sheet.getNumMergedRegions();

        int row = getrow(cell)-1;
        int column = getcolumn(cell)-1;
        //System.out.println("Check,"+row+","+column+",合并的单元格数:"+sheetMergeCount);
        boolean result = false;

        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            //System.out.println("result:"+result+","+firstRow+","+lastRow+","+firstColumn+","+lastColumn);//上，下，左，右
            if(row >= firstRow && row <= lastRow){
                if(column >= firstColumn && column <= lastColumn){
                    result |= true;
                }
            } else {
                result |= false;
            }
        }
        return result;
    }

    /**
     *
     * @param sheetname 表名
     * @param cell 单元格
     * @return 返回簇状图????????????????????????????????????????????????????????????????????
     */
    public boolean getChart(String sheetname,String cell){
        return false;
    }

    /**
     *返回单元格数值格式
     * @param sheetname 表名
     * @param cell 单元格
     * @return 返回单元格数值格式
     */
    public String getFormat(String sheetname,String cell){
        return workBook.getSheet(sheetname).getRow(getrow(cell)-1).getCell(getcolumn(cell)-1).getCellStyle().getDataFormatString();
    }

    /**
     *返回单元格公式
     * @param sheetname 表名
     * @param cell 单元格
     * @return 返回单元格公式
     */
    public String getCellFormula(String sheetname,String cell){
        return workBook.getSheet(sheetname).getRow(getrow(cell)-1).getCell(getcolumn(cell)-1).getCellFormula();
    }

    /**
     *返回单元格名称
     * @param sheetname 表名
     * @param value 单元格名称
     * @return 返回单元格名称
     */
    public boolean getCellName(String sheetname,String value){
        try {
            String sheetname1 = workBook.getSheet(workBook.getName(value).getSheetName()).getSheetName();
            if(sheetname.equals(sheetname1)){
                return true;
            }
            else {
                return false;
            }
        }catch(Exception ex){
            return false;
        }
    }

    /**
     *返回对齐方式
     * @param sheetname 表名
     * @param cell 单元格
     * @return 返回对齐方式
     */
    public String getCellAlignment(String sheetname,String cell){
        return workBook.getSheet(sheetname).getRow(getrow(cell)-1).getCell(getcolumn(cell)-1).getCellStyle().getAlignment().name();
    }

    /**
     *返回是否存在数据透视表
     * @param sheetname 表名
     * @return 返回是否存在数据透视表
     */
    public boolean isPivotTable(String sheetname){
        try{
            for(XSSFPivotTable xssfPivotTable :workBook.getSheet(sheetname).getPivotTables()){
                return true;
            }}catch (Exception e){
            return false;
        }
        return false;
    }

    /**
     *返回是否存在边框线(上下左右)
     * @param sheetname 表名
     * @return 返回是否存在边框线
     */
    public boolean isBorderLineBottom(String sheetname,String cell){
        try{
            short sBut = workBook.getSheet(sheetname).getRow(getrow(cell)-1).getCell(getcolumn(cell)-1).getCellStyle().getBorderBottom().getCode();
            //System.out.println("sBut="+sBut);
            if(sBut>0) {
                return true;
            } else {
                return false;
            }
        }catch (Exception e){
            return false;
        }
    }
    public boolean isBorderLineLeft(String sheetname,String cell){
        try{
            short sLut = workBook.getSheet(sheetname).getRow(getrow(cell)-1).getCell(getcolumn(cell)-1).getCellStyle().getBorderLeft().getCode();
            //System.out.println("sLut="+sLut);
            if(sLut>0) {
                return true;
            } else {
                return false;
            }
        }catch (Exception e){
            return false;
        }
    }
    public boolean isBorderLineRight(String sheetname,String cell){
        try{
            short sRut = workBook.getSheet(sheetname).getRow(getrow(cell)-1).getCell(getcolumn(cell)-1).getCellStyle().getBorderRight().getCode();
            //System.out.println("sRut="+sRut);
            if(sRut>0) {
                return true;
            } else {
                return false;
            }
        }catch (Exception e){
            return false;
        }
    }
    public boolean isBorderLineTop(String sheetname,String cell){
        try{
            short sTut = workBook.getSheet(sheetname).getRow(getrow(cell)-1).getCell(getcolumn(cell)-1).getCellStyle().getBorderTop().getCode();
            //System.out.println("sTut="+sTut);
            if(sTut>0) {
                return true;
            } else {
                return false;
            }
        }catch (Exception e){
            return false;
        }
    }

    /**
     *返回行高
     * @param sheetname 表名
     * @param cell 单元格
     * @return 返回行高,不是1/20！！！！！
     */
    public short getCellHeight(String sheetname,String cell){
        return  workBook.getSheet(sheetname).getRow(getrow(cell)-1).getHeight();
    }

    /**
     *返回列宽，是1/256个字符宽度//in units of 1/256th of a character width )
     * @param sheetname 表名
     * @param cell 单元格
     * @return 返回列宽,是1/256个字符宽度//in units of 1/256th of a character width )
     */
    public int getCellWidth(String sheetname,String cell){
        return  workBook.getSheet(sheetname).getColumnWidth(getcolumn(cell)-1);
    }

    /**
     *返回List<String>,例如"A3:B6"->"A3","A4","A5","A6","B3","B4","B5","B6";
     * @param cellsarea 单元格区域
     * @return 返回List<String>,例如"A3:B6"->"A3","A4","A5","A6","B3","B4","B5","B6";
     *
     * */
    public List<String> getCellsList(String cellsarea){
        String top = cellsarea.substring(0, cellsarea.indexOf(":"));
        String buttom = cellsarea.substring(cellsarea.indexOf(":")+1,cellsarea.length());
        CellRangeAddress cellAddresses = new CellRangeAddress( getrow(top),getrow(buttom),getcolumn(top),getcolumn(buttom));
        List<String> listcells = new ArrayList<String>();
        for(int i=getrow(top);i<=getrow(buttom);i++){
            for(int j = (getcolumn(top)-1);j<=(getcolumn(buttom)-1);j++){
                char letter = (char)((int)'A'+j);
                listcells.add(letter+""+i);
            }
        }
        return listcells;
    }

    /**
     *把cell（A3）转化为(13)
     * @param cell 把cell（A3）转化为(13)
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
