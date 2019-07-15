package com.gzmut.office.util;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;

import java.awt.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DecimalFormat;

/**
 * 对演示文稿的处理工具
 * @author livejq
 * @date 2019/7/13
 **/
public class PptUtils {

    /**
     * 获取ppt演示文档对象实例
     * @param fileName ppt演示文档文件路径
     * @return org.apache.poi.xslf.usermodel.XMLSlideShow
     **/
    public static XMLSlideShow getXMLSlideShow(String fileName) {

        if(fileName == null || fileName.equals("")) {
            return null;
        }
        if(!fileName.endsWith(".ppt") || !fileName.endsWith(".pptx")) {
            // 最好抛出异常提示信息
            return null;
        }
        XMLSlideShow slideShow = null;
        // 读取ppt演示文档
        try {
            slideShow = new XMLSlideShow(new FileInputStream(fileName));
        } catch (IOException e) {
            e.printStackTrace();
        }

        return slideShow;
    }

    /**
     * 获取幻灯片总数
     * @param slideShow ppt演示文档对象
     * @return int
     */
    public static int getSlidesSum(XMLSlideShow slideShow) {
        return slideShow.getSlides().size();
    }

    /**
     * 获取ppt演示文档中的幻灯片尺寸
     * @param slideShow ppt演示文档对象
     * @return java.awt.Dimension
     **/
    public static Dimension getSlideSize(XMLSlideShow slideShow) {
        return slideShow.getPageSize();
    }

    /**
     * 获取幻灯片板式
     * @param slide 幻灯片对象
     * @return org.apache.poi.xslf.usermodel.SlideLayout
     **/
    public static SlideLayout getSlideLayoutType(XSLFSlide slide) {
        return slide.getSlideLayout().getType();
    }

    /**
     * 获取幻灯片标题
     * @param slide 幻灯片对象
     * @return java.lang.String
     **/
    public static String getSlideTitle(XSLFSlide slide) {
        return slide.getTitle();
    }

    /**
     * 获取幻灯片编号
     * @param slide 幻灯片对象
     * @return int
     **/
    public static int getSlideId(XSLFSlide slide) {
        return slide.getSlideNumber();
    }

    /**
     * 获取幻灯片主题名称
     * @param slide 幻灯片对象
     * @return java.lang.String
     **/
    public static String getSlideTheme(XSLFSlide slide) {
        return slide.getTheme().getName();
    }

    /**
     * 获取幻灯片填充颜色
     * @param slide 幻灯片对象
     * @return java.awt.Color
     **/
    public static Color getSlideFillColor(XSLFSlide slide) {
        return slide.getBackground().getFillColor();
    }

    /**
     * 获取幻灯片填充颜色的透明度
     * @param slide 幻灯片对象
     * @return int
     **/
    public static int getSlideFillColorAlpha(XSLFSlide slide) {
        return slide.getBackground().getFillColor().getAlpha();
    }

    /**
     * 获取某个文本框填充的颜色
     * @param slide 幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @return java.awt.Color
     **/
    public static Color getTextBoxFillColor(XSLFSlide slide, int TextBoxIndex) {
        return slide.getPlaceholder(TextBoxIndex).getFillColor();
    }

    /**
     * 获取某个文本框中的文字内容
     * @param slide 幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @return java.lang.String
     **/
    public static String getTextBoxContent(XSLFSlide slide, int TextBoxIndex) {
        return slide.getPlaceholder(TextBoxIndex).getText();
    }

    /**
     * 获取某个文本框中的段落总数
     * @param slide 幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @return int
     **/
    public static int getTextBoxParagraphSum(XSLFSlide slide, int TextBoxIndex) {
        return slide.getPlaceholder(TextBoxIndex).getTextBody().getParagraphs().size();
    }

    /**
     * 获取某个文本框中的行高（存在单位换算问题<磅>）
     * @param slide 幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @return int
     **/
    public static double getTextBoxLineHeight(XSLFSlide slide, int TextBoxIndex) {
        return slide.getPlaceholder(TextBoxIndex).getTextHeight();
    }

    /**
     * 获取某个文本框中的某一段内容
     * @param slide 幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @return java.lang.String
     **/
    public static String getTextBoxParagraph(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex) {
        return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getText();
    }

    /**
     * 获取某个文本框中某一段的段前间距（存在单位换算问题<磅>）
     * @param slide 幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @return java.lang.Double
     **/
    public static Double getTextBoxParagraphSpaceBefore(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex) {
        return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getSpaceBefore();
    }

    /**
     * 获取某个文本框中某一段的段后间距（存在单位换算问题<磅>）
     * @param slide
     * @param TextBoxIndex
     * @param ParagraphIndex
     * @return java.lang.Double
     **/
    public static Double getTextBoxParagraphSpaceAfter(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex) {
        return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getSpaceAfter();
    }

    /**
     * 获取某个文本框中某一段的行距（存在单位换算问题<磅>）
     * @param slide
     * @param TextBoxIndex
     * @param ParagraphIndex
     * @return java.lang.Double
     **/
    public static Double getTextBoxParagraphLineSpace(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex) {
        return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getLineSpacing();
    }

    /**
     * 获取某个文本框中某一段内容中的某段字符串
     * @param slide 幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @param RawIndex 字符串编号（从0开始，类型区别主要为数字、英文字母、中文等）
     * @return java.lang.String
     **/
    public static String getTextBoxRawContent(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex, int RawIndex) {
        return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getTextRuns().get(RawIndex).getRawText();
    }

    /**
     * 获取某个文本框中某一段内容中的某段字符串中的超链接
     * @param slide 幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @param RawIndex 字符串编号（从0开始，类型区别主要为数字、英文字母、中文等）
     * @return java.lang.String
     **/
    public static String getTextBoxRawHyperlink(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex, int RawIndex) {
        return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getTextRuns().get(RawIndex).getHyperlink().getAddress();
    }

    /**
     * 获取某个文本框中某一段内容中的某段字符串的字体大小
     * @param slide 幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @param RawIndex 字符串编号（从0开始，类型区别主要为数字、英文字母、中文等）
     * @return int
     **/
    public static int getTextBoxRawFontSize(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex, int RawIndex) {
        return Integer.parseInt(new DecimalFormat("##").format(slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getTextRuns().get(RawIndex).getFontSize()));
    }

    /**
     * 获取某个文本框中某一段内容中的某段字符串的字体样式
     * @param slide 幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @param RawIndex 字符串编号（从0开始，类型区别主要为数字、英文字母、中文等）
     * @return java.lang.String
     **/
    public static String getTextBoxRawFontFamily(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex, int RawIndex) {
        return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getTextRuns().get(RawIndex).getFontFamily();
    }

    /**
     * 判断某个文本框中某一段内容中的某段字符串是否粗体
     * @param slide 幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @param RawIndex 字符串编号（从0开始，类型区别主要为数字、英文字母、中文等）
     * @return boolean
     **/
    public static boolean isTextBoxRawBold(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex, int RawIndex) {
        return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getTextRuns().get(RawIndex).isBold();
    }

    /**
     * 判断某个文本框中某一段内容中的某段字符串是否存在脚注
     * @param slide 幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @param RawIndex 字符串编号（从0开始，类型区别主要为数字、英文字母、中文等）
     * @return boolean
     **/
    public static boolean isTextBoxRawSubscript(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex, int RawIndex) {
        return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getTextRuns().get(RawIndex).isSubscript();
    }

    /**
     * 判断某个文本框中某一段内容中的某段字符串是否存在下划线
     * @param slide 幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @param RawIndex 字符串编号（从0开始，类型区别主要为数字、英文字母、中文等）
     * @return boolean
     **/
    public static boolean isTextBoxRawUnderlined(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex, int RawIndex) {
        return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getTextRuns().get(RawIndex).isUnderlined();
    }
}
