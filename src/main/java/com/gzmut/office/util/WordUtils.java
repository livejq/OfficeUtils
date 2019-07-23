package com.gzmut.office.util;

import com.alibaba.fastjson.JSONObject;
import com.beust.jcommander.internal.Lists;
import com.google.common.base.CharMatcher;
import com.gzmut.office.enums.word.*;
import com.gzmut.office.service.WordCorrect;
import lombok.extern.slf4j.Slf4j;
import opennlp.tools.util.wordvector.WordVectorTable;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.w3c.dom.NamedNodeMap;

import java.io.*;

import java.math.BigInteger;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static com.gzmut.office.enums.word.WordBackgroundPropertiesEnums.*;
import static com.gzmut.office.enums.word.WordCorrectEnums.CHECK_PAGE;

/**
 * Word文档的工具类
 * @author MXDC
 * @date 2019/7/9
 **/
@Slf4j
public class WordUtils {

    /** word文档的文档文件名*/
    public static  String FILE_NAME="";
    /** 考试文件夹**/
    public static String examDir = "";
    /** word文档 XWPFPDocumet类*/
    public static  XWPFDocument document;

    public static final String DOC_SUFFIX = ".doc";

    public static final String DOCX_SUFFIX = ".docx";
    /**
     * 给单前工具类设置一个word文档
     * @param fileName word文档 全路径名
     */
    public static void setDocment(String fileName){
        // 文件名为空或者null,返回null
        if (fileName == null || "".equals(fileName)){
            return ;
        }
        // 设置文件名
        FILE_NAME = fileName;
        InputStream stream = null;
        try {
            stream = new FileInputStream(new File(fileName));
            //word为2003，进行转换
            if (fileName.endsWith(DOC_SUFFIX)){
                //进行转换
                //document=;
            } else if (fileName.endsWith(DOCX_SUFFIX)){
                document = new XWPFDocument(stream);
            } else {
                log.error(fileName+"不是word文件");
            }
        }catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (null != stream){
                try {
                    stream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                    log.error("读取word文件失败");
                }
            }
        }
        return ;

    }
    /**
     *  获取word文档的XWPFDocument类，未完善.doc的转换
     * @param fileName 文件名
     * @return
     * @throws IOException
     */
    public static XWPFDocument getDocument(String fileName)  {
        XWPFDocument document = null;
        // 文件名为空或者null,返回null
        if (fileName == null || "".equals(fileName)){
            return null;
        }
        // 获取文件名
        FILE_NAME = fileName;
        InputStream stream = null;
        try {
            stream = new FileInputStream(new File(fileName));
            //word为2003，进行转换
            if (fileName.endsWith(DOC_SUFFIX)){
                //进行转换
                //document=;
            } else if (fileName.endsWith(DOCX_SUFFIX)){
                document = new XWPFDocument(stream);
            } else {
                log.error(fileName+"不是word文件");
            }
        }catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (null != stream){
                try {
                    stream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                    log.error("读取word文件失败");
                }
            }
        }
		return document;
    }

    /**
     * 获取word文档的纯文本
     * @param fileName
     * @param <T>
     * @return

    public static <T> List<String> readWordFile(String fileName) {
        // 文件名为空或者null,返回null
        if (fileName == null || "".equals(fileName)){
            return null;
        }
        List<String> contextList = Lists.newArrayList();
        InputStream stream = null;
        try {
            stream = new FileInputStream(new File(fileName));
            if (fileName.endsWith(DOC_SUFFIX)) {
                HWPFDocument document = new HWPFDocument(stream);
                WordExtractor extractor = new WordExtractor(document);
                String[] contextArray = extractor.getParagraphText();
                Arrays.asList(contextArray).forEach(context -> contextList.add(CharMatcher.whitespace().removeFrom(context)));
                extractor.close();
                document.close();
            } else if (fileName.endsWith(DOC_SUFFIX)) {
                XWPFDocument document = new XWPFDocument(stream).getXWPFDocument();
                List<XWPFParagraph> paragraphList = document.getParagraphs();
                paragraphList.forEach(paragraph -> contextList.add(CharMatcher.whitespace().removeFrom(paragraph.getParagraphText())));
                document.close();
            } else {
                log.error(fileName+"不是word文件");
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (null != stream){
                try {
                    stream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                    log.error("读取word文件失败");
                }
            }
        }
        return contextList;
    }*/

    /**
     * 通过字符串获取段落
     * @param s 所在段落的字符串
     * @return XWPFParagraph word 段落对象
     */
    public static XWPFParagraph getParagraph(String s){
        // 获取文档的所有段落
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        // 获取文本所在段落
        XWPFParagraph xwpfParagraph = paragraphs.stream().filter(paragraph -> paragraph.getText().contains(s))
                .findFirst().get();
        return xwpfParagraph;
    }
    /**
     * 在段落里查找字符串
     * 返回一个字符串Run
     * @param paragraph 一个段落
     * @param s 要查找的字符串
     * @return XWPFRUN
     */
    public static XWPFRun getRun(XWPFParagraph paragraph, String s){
        return paragraph.getRuns().stream().filter(xwpfRun -> xwpfRun.getText(0).contains(s)).findFirst().get();
    }

    /**
     * 获取段落属性
     * @param paragraph 所在段落
     * @param property  属性
     * @return
     */
    public static String getParagraphProperties(XWPFParagraph paragraph, String property){
        WordParagraphPropertiesEnums wordParagraphPropertiesEnum = null;
        if( property ==null || "".equals(property)){
            return null;
        }
        try {
            // 将String转化为Enum
            wordParagraphPropertiesEnum = WordParagraphPropertiesEnums.valueOf(property.toUpperCase());
        }catch (IllegalArgumentException e){
            e.printStackTrace();
            return null;
        }
        switch (wordParagraphPropertiesEnum){
            case ALIGNMENT: return paragraph.getAlignment().name();
            default:;
        }
        return null;
    }

    /**
     * 获取页面背景颜色
     * 返回16进制颜色值
     * @return String
     */
    public static String getBackgroundColor(){
        CTDocument1 ctDocument1 = document.getDocument();
        CTBackground background = ctDocument1.getBackground();
        if (background == null){
            return  null;
        }
        STHexColor stHexColor = background.xgetColor();
        return stHexColor == null ? null : stHexColor.getStringValue();
    }

    /**
     * 获取背影填充效果属性
     * @return NameNodeMap
     */
    public static NamedNodeMap getBackgroundFillAttributes(){
        CTDocument1 document = WordUtils.document.getDocument();
        CTBackground background = document.getBackground();
        if (background == null){
            return  null;
        }
        //通过 xpath选取v:fill节点
        XmlObject[] xmlObjects = background.selectPath("declare namespace v='urn:schemas-microsoft-com:vml' " + ".//v:fill");
        if( xmlObjects.length == 0){
            return null;
        }
        // 获取节点属性
        NamedNodeMap attributes = xmlObjects[0].getDomNode().getAttributes();
        return attributes;
    }

    /**
     * 获取背景-填充效果-纹理
     * @return String
     */
    public  static String getBackgroundFillTile(){
        NamedNodeMap attributes = getBackgroundFillAttributes();
        if (attributes == null){
            return null;
        }
        String type = attributes.getNamedItem("type").getNodeValue();
        //判断填充效果是否为纹理
        if (!"tile".equals(type)){
            return null;
        }
        // 获取填充效果纹理名称
        String nodeValue = attributes.getNamedItem("o:title").getNodeValue();
        return nodeValue;
    }

    /**
     * 获取背景-填充效果-图案
     * @return String
     */
    public static String getBackgroundFillPattern(){
        NamedNodeMap attributes = getBackgroundFillAttributes();
        if (attributes == null){
            return null;
        }
        String type = attributes.getNamedItem("type").getNodeValue();
        //判断填充效果是否为图案
        if (!"pattern".equals(type)){
            return null;
        }
        // 获取填充效果图案名称
        String nodeValue = attributes.getNamedItem("o:title").getNodeValue();

        return nodeValue;
    }

    /**
     * 检测背景颜色是否一致
     * @param color 背景16进制颜色
     * @return boolean
     */
    public static boolean checkBackgroundColor(String color){
        if(StringUtils.isEmpty(color)){
            return false;
        }
        return color.equals(getBackgroundColor()) ? true:false;
    }

    /**
     * 检查背景是否正确
     * @param property 背景的属性
     * @param value 背景属性的值
     * @return
     */
    public static boolean checkBackground(String property, String value){
        WordBackgroundPropertiesEnums wordBackgroundPropertiesEnums = null;
        if( property ==null || "".equals(property)){
            return false;
        }
        try {
              // 将String转化为Enum
              wordBackgroundPropertiesEnums = WordBackgroundPropertiesEnums.valueOf(property.toUpperCase());
        }catch (IllegalArgumentException e){
            e.printStackTrace();
            return false;
        }
        // 属性检查
        switch (wordBackgroundPropertiesEnums){
            case BACKGROUND_COLOR:
                break;
            default:
        }
        return false;
    }

    /**
     * 全文范围内搜索，如果存在返回true
     * @param text 要搜索的字符串
     * @return 正确或错误信息
     */
    public static String checkAllText(String text){

        if(document == null){
            System.out.println("document为空");
            return  null;
        }
        List<XWPFParagraph> list = document.getParagraphs();
        for(XWPFParagraph para : list){
            String t = para.getText();
            if(t.contains(text)){
                return "true";
            }
        }
        return "false:文档不存在文本"+text;
    }

    /**
     * @return
     */
    public String checkDiffFont(){
        List<XWPFParagraph> list = document.getParagraphs();
        String fName = "";
        List<XWPFRun> runs = list.get(0).getRuns();
        for(XWPFRun r : runs){
            fName = r.getFontName();
        }
        for(XWPFParagraph para : list) {
            runs = para.getRuns();
            for(XWPFRun r : runs){
                if(!fName.equals(r.getFontName())) {
                    return "true";
                }
            }

        }
        return "false:字体相同";
    }

    /**
     * 判断文件是否存在
     * 需要参数 fileName
     * @param param 判断参数
     * @return Map<Stirng,Object> 判题信息
     */
    public static Map<String,Object> checkFileIsExist(int score, Map<String, Object> param){
        Map<String, Object> infoMap = new HashMap<>(2);
        if (new File(examDir+File.separator+param.get("fileName").toString()).exists()){
            infoMap.put("score",score);
            infoMap.put("msg","文件存在,答题正确");
        }else {
            infoMap.put("score",0);
            infoMap.put("msg","文件不存在");
        }
        return infoMap;
    }

    /**
     * 检查页面大小
     * @param pageSize 页面大小 参数格式 ：w:h
     * @return
     */
    public static boolean checkPageSize(String pageSize){
        if (StringUtils.isEmpty(pageSize)){
            return false;
        }
        pageSize = pageSize.trim();
        String w = StringUtils.substringBefore(pageSize,":");
        String h = StringUtils.substringAfter(pageSize,":");
        CTDocument1 document = WordUtils.document.getDocument();
        CTPageSz pgSz = document.getBody().getSectPr().getPgSz();
        if ( w.equals(pgSz.getW().toString()) && h.equals(pgSz.getH().toString())){
            return true;
        }
        return false;
    }

    /**
     * 判断页边距是否正确
     * 满足参数中的页边距都真确才得分
     * @param param
     *
     * @return boolena
     */
    /**
     *  判断页边距是否正确（包括页眉，页脚距边界）
     * @param key 参数-键
     * @param value 参数-值
     * @return
     */
    public static boolean checkPageMargin(String key, Object value){
        CTDocument1 document = WordUtils.document.getDocument();
        CTPageMar pgMar = document.getBody().getSectPr().getPgMar();
        long top = pgMar.getTop().longValue();
        long bottom = pgMar.getBottom().longValue();
        long right = pgMar.getRight().longValue();
        long left = pgMar.getLeft().longValue();
        long header = pgMar.getHeader().longValue();
        long footer = pgMar.getFooter().longValue();
        long val = Long.valueOf(value.toString());
        WordPagePropertiesEnums wordPagePropertiesEnums = WordPagePropertiesEnums.valueOf(key);
        switch (wordPagePropertiesEnums){
            case PAGE_MARGIN_TOP:
                return val == top ? true:false;
            case PAGE_MARGIN_BOTTOM:
               return val == bottom ? true : false;
            case PAGE_MARGIN_RIGHT:
                return val == right ? true : false;
            case PAGE_MARGIN_LEFT:
                return val == left ? true : false;
            case PAGE_MARGIN_HEADER:
                return val == header ? true : false;
            case PAGE_MARGIN_FOOTER:
                return val == footer ? true : false;
            default:
        }
        return true;
    }

    /**
     * 检查页面参数
     * @param score 分数值
     * @param param 参数
     * @return Map<Stirng,Object> 判题信息
     */
    public static Map<String,Object> checkPage(int score, Map<String, Object> param){
        Map<String, Object> infoMap = new HashMap<>(2);
        StringBuffer msg = new StringBuffer();
        int setpScore =  score/param.size();
        int tatalScore = 0;
        for (String key: param.keySet()) {
            WordPagePropertiesEnums pagePropertiesEnums = WordPagePropertiesEnums.valueOf(key);
            switch (pagePropertiesEnums){
                case PAGE_SIZE:
                    if (checkPageSize(param.get(key).toString())){
                        tatalScore += setpScore;
                        msg.append("页面大小正确;");
                    }else {
                        msg.append("页面大小错误;");
                    }
                    break;
                case PAGE_MARGIN_TOP:
                case PAGE_MARGIN_BOTTOM:
                case PAGE_MARGIN_LEFT:
                case PAGE_MARGIN_RIGHT:
                case PAGE_MARGIN_HEADER:
                case PAGE_MARGIN_FOOTER:
                    if (checkPageMargin(key, param.get(key))) {
                        tatalScore += setpScore;
                        msg.append(pagePropertiesEnums.getProperty()+"正确;");
                    } else {
                        msg.append(pagePropertiesEnums.getProperty()+"不正确;");
                    }
                    break;
                default:
            }
        }
        infoMap.put("score",tatalScore);
        infoMap.put("msg",msg.toString());
        log.info(infoMap.toString());
        return  infoMap;
    }

    public static boolean checkTableProperties(String key, int value){
        List<XWPFTable> tables = WordUtils.document.getTables();
        XWPFTable table = tables.get(0);
        int result = 0;
        WordTablePropertiesEnums enums = WordTablePropertiesEnums.valueOf(key);
        switch(enums){
            case TABLE_WIDTH:
                List<XWPFTableCell> cells = table.getRow(0).getTableCells();
                for(XWPFTableCell cell : cells){
                    result += cell.getWidth();
                }
                return GeneralUtils.approach(result, value);
            case TABLE_HEIGHT:
                List<XWPFTableRow> rows = table.getRows();
                for(XWPFTableRow row : rows){
                    result += row.getHeight();
                }
                return GeneralUtils.approach(result, value);
            default:
                return false;
        }

    }
}

