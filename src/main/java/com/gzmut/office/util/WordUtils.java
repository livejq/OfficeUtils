package com.gzmut.office.util;

import com.beust.jcommander.internal.Lists;
import com.google.common.base.CharMatcher;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.util.Arrays;
import java.util.List;

/**
 * Word文档的工具类
 * @author MXDC
 * @date 2019/7/9
 **/
public class WordUtils {

    /** word文档的文档文件名*/
    public static  String FILE_NAME="";
    /** word文档 XWPFPDocumet类*/
    public static  XWPFDocument document;

    /**
     *  获取word文档的XWPFDocument类，未完善.doc的转换
     * @param fileName 文件名
     * @return
     * @throws IOException
     */
    public XWPFDocument getXWPDocument(String fileName)  {
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
            if (fileName.endsWith(".doc")){
                //进行转换
                //document=;
            } else if (fileName.endsWith(".docx")){
                System.out.println(fileName);
                document = new XWPFDocument(stream);
            } else {
//                LOGGER.debug("此文件{}不是word文件", path);
                //这里要给为logger
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
//                    LOGGER.debug("读取word文件失败");
                    System.out.printf("读取word文件失败");
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
     */
    public static <T> List<String> readWordFile(String fileName) {
        // 文件名为空或者null,返回null
        if (fileName == null || "".equals(fileName)){
            return null;
        }
        List<String> contextList = Lists.newArrayList();
        InputStream stream = null;
        try {
            stream = new FileInputStream(new File(fileName));
            if (fileName.endsWith(".doc")) {
                HWPFDocument document = new HWPFDocument(stream);
                WordExtractor extractor = new WordExtractor(document);
                String[] contextArray = extractor.getParagraphText();
                Arrays.asList(contextArray).forEach(context -> contextList.add(CharMatcher.whitespace().removeFrom(context)));
                extractor.close();
                document.close();
            } else if (fileName.endsWith(".docx")) {
                XWPFDocument document = new XWPFDocument(stream).getXWPFDocument();
                List<XWPFParagraph> paragraphList = document.getParagraphs();
                paragraphList.forEach(paragraph -> contextList.add(CharMatcher.whitespace().removeFrom(paragraph.getParagraphText())));
                document.close();
            } else {
//                LOGGER.debug("此文件{}不是word文件", path);
                //这里要给为logger
                System.out.printf("此文件不是word文件");
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (null != stream){
                try {
                    stream.close();
                } catch (IOException e) {
                    e.printStackTrace();
//                    LOGGER.debug("读取word文件失败");
                    System.out.printf("读取word文件失败");
                }
            }
        }
        return contextList;
    }

    /**
     * 全文范围内搜索，如果存在返回true
     * @param text 要搜索的字符串
     * @return 正确或错误信息
     */
    public String checkAllText(String text){
        //stream = new FileInputStream(new File(fileName));
        //XWPFDocument document = new XWPFDocument(stream).getXWPFDocument();
        //List<XWPFParagraph> paragraphList = document.getParagraphs();
        List<XWPFParagraph> list = document.getParagraphs();
        for(XWPFParagraph para : list){
            String t = para.getText();
            if(t.contains(text)){
                return "true";
            }
        }
        return "false:文档不存在文本"+text;
    }

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
                if(!fName.equals(r.getFontName())) return "true";
            }

        }
        return "false:字体相同";
    }
}
