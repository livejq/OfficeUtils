package com.gzmut.office.util;

import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;

import java.util.List;

import static org.junit.Assert.*;

public class WordUtilsTest {
    @Test
    public void Test(){
        WordUtils wordUtils = new WordUtils();
        XWPFDocument document = wordUtils.getXWPDocument("F:\\KSWJJ\\15000001\\WORD1.docx");
       // List<String> strings = WordUtils.readWordFile("E:\\Desktop\\java判断word文档字体.docx");
        //strings.forEach(string->System.out.println(string));

        List<XWPFParagraph> paragraphs = document.getParagraphs();
//        paragraphs.forEach(xwpfParagraph -> System.out.println(xwpfParagraph));
//        paragraphs.stream().forEach(p->System.out.println(p.getText()));
        //获取文本所在段落
        XWPFParagraph xwpfParagraph = paragraphs.stream().filter(paragraph -> paragraph.getText().contains("调查表明京沪穗网民主导“B2C”"))
                .findFirst().get();
        //获取段落文本,包括图片文本
        System.out.printf(xwpfParagraph.getText());
        //不包括图片文本
        System.out.printf(xwpfParagraph.getParagraphText());
        // @return the paragraph alignment of this paragraph.
//                LEFT(1),
//                CENTER(2),
//                RIGHT(3),
//                BOTH(4),
//                MEDIUM_KASHIDA(5),
//                DISTRIBUTE(6),
//                NUM_TAB(7),
//                HIGH_KASHIDA(8),
//                LOW_KASHIDA(9),
//                THAI_DISTRIBUTE(10);
        System.out.print(xwpfParagraph.getAlignment().getValue());
        xwpfParagraph.getVerticalAlignment();

        /**
         * 首行缩进
         * Specifies the additional indentation which shall be applied to the first
         * line of the parent paragraph. This additional indentation is specified
         * relative to the paragraph indentation which is specified for all other
         * lines in the parent paragraph.
         * The firstLine and hanging attributes are
         * mutually exclusive, if both are specified, then the firstLine value is
         * ignored.
         * If the firstLineChars attribute is also specified, then this
         * value is ignored.
         * If this attribute is omitted, then its value shall be
         * assumed to be zero (if needed).
         *
         * @return indentation or null if indentation is not set
         */
        System.out.print( xwpfParagraph.getFirstLineIndent());
        //缩进
        xwpfParagraph.getIndentationLeft();
        xwpfParagraph.getIndentationRight();
        xwpfParagraph.getIndentationFirstLine();
        //行间距???
        System.out.println(xwpfParagraph.getSpacingLineRule().getValue());
        //段前段后距
        System.out.println(xwpfParagraph.getSpacingAfter());
        System.out.println(xwpfParagraph.getSpacingAfterLines());
        //段落编号格式
        xwpfParagraph.getNumFmt();
        //字符串
        XWPFRun xwpfRun = xwpfParagraph.getRuns().stream().filter(run -> run.getText(0).contains("调"))
                .findFirst().get();
        xwpfRun.getFontSize();
        xwpfRun.getFontName();
        xwpfRun.getFontFamily();
        xwpfRun.getVerticalAlignment();
        xwpfRun.getUnderline();
        xwpfRun.getUnderlineColor();
        xwpfRun.getColor();
        xwpfRun.isBold();
        xwpfRun.isItalic();
        xwpfRun.isShadowed();
        xwpfRun.isCapitalized();
        xwpfRun.isDoubleStrikeThrough();

//        System.out.println(xwpfParagraph.getRuns().size());
////        xwpfParagraph.getRuns().stream().filter(run->run.getText(0)!);
////        System.out.println(xwpfRun.getText(0));
//        xwpfParagraph.getRuns().forEach(xwpfRun -> System.out.println(xwpfRun.getText(0)));

        //获取表格
        List<XWPFTable> tables = document.getTables();

        //获取图片
        List<XWPFPictureData> allPictures = document.getAllPictures();
        allPictures.forEach(a->a.getRelations()
        );
//        document.getAllPackagePictures();

    }
}