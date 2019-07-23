package com.gzmut.office.util;

import com.gzmut.office.enums.word.WordBackgroundPropertiesEnums;
import com.microsoft.schemas.vml.CTShape;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.SimpleValue;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;

import org.junit.Test;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;

import javax.xml.namespace.QName;
import java.io.IOException;
import java.util.List;
import java.util.Scanner;
import java.util.stream.Stream;

import static org.junit.Assert.*;

@Slf4j
public class WordUtilsTest {

   

    @Test
    public void getWordParagraphProperties(){
        WordUtils.setDocment("F:\\KSWJJ\\15000001\\WORD1.docx");
        XWPFDocument document = WordUtils.document;
        XWPFParagraph paragraph = WordUtils.getParagraph("调查表明京沪");
        System.out.println(WordUtils.getParagraphProperties(paragraph,"alignment"));

    }

    @Test
    public void getWordBackgroundFillTest(){
        WordUtils.setDocment("F:\\KSWJJ\\15000001\\WORD1.docx");
        // 获取背景填充纹理
        String backgroundFillTile = WordUtils.getBackgroundFillTile();
        System.out.println(backgroundFillTile);
        // 获取背景填充图案
        String backgroundFillPattern = WordUtils.getBackgroundFillPattern();
        System.out.println(backgroundFillPattern);
    }

    @Test
    public void Test1(){
        WordUtils.setDocment("F:\\KSWJJ\\15000001\\WORD1.docx");
        XWPFDocument document = WordUtils.document;
//      System.out.println(document.getDocument().getBackground());
//      System.out.println(WordUtils.getBackgroundColor());
        CTBackground background = document.getDocument().getBackground();
        System.out.println(background);
        XmlObject[] xmlObjects = background.selectPath("declare namespace v='urn:schemas-microsoft-com:vml' " + ".//v:fill");
        NamedNodeMap attributes = xmlObjects[0].getDomNode().getAttributes();
        System.out.println(attributes.getNamedItem("type").getNodeValue());
//        System.out.println(attributes.getNamedItem("o:title").getNodeValue());
//        System.out.println(xmlObjects[0].getDomNode().getNodeName());
//        System.out.println(background.selectChildren(new QName("urn:schemas-microsoft-com:vml", "fill", "v")).length);
//        System.out.println(background.selectPath("declare namespace v='urn:schemas-microsoft-com:vml' " + ".//v:fill").length);
//        System.out.println(background.getDomNode().getFirstChild().getNamespaceURI());
//        System.out.println(background.getDomNode().getFirstChild().getNodeName());


        // 根据关键字获取段落
//        XWPFParagraph paragraph = WordUtils.getParagraph("调查还发现");
//        paragraph.getRuns().forEach(xwpfRun -> System.out.println(xwpfRun.getText(0)));
//        WordBackgroundPropertiesEnums background_color = WordBackgroundPropertiesEnums.valueOf("ackground_color".toUpperCase());



    }

    @Test
    public void Test(){

        XWPFDocument document = WordUtils.getDocument("F:\\KSWJJ\\15000001\\WORD1.docx");
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
        /*
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

    @Test
    public void pageSize(){
        WordUtils.setDocment("F:\\KSWJJ\\15000001\\WORD1.docx");
        XWPFDocument document = WordUtils.document;
        CTDocument1 document1 = document.getDocument();
        CTSectPr sectPr = document1.getBody().getSectPr();
        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz pgSz = sectPr.getPgSz();
        System.out.println(pgSz);
        System.out.println(pgSz.getW());
        System.out.println(pgSz.getH());
    }

    @Test
   public void pageMargin(){
       WordUtils.setDocment("F:\\KSWJJ\\15000001\\WORD1.docx");
       XWPFDocument document = WordUtils.document;
       CTDocument1 document1 = document.getDocument();
       CTPageMar pgMar = document1.getBody().getSectPr().getPgMar();
       System.out.println(pgMar.getBottom().toString());
       System.out.println(pgMar.getTop());
       System.out.println(pgMar.getLeft());
       System.out.println(pgMar.getRight());
        System.out.println(pgMar.getHeader());
        System.out.println(pgMar.getFooter());

    }

   @Test
   public void haeder(){

       WordUtils.setDocment("C:\\KSWJJ\\65910001\\参考答案\\Word.docx");
       XWPFDocument document = WordUtils.document;
       // 获取页眉引用列表
       CTSectPr sectPr = document.getDocument().getBody().getSectPr();
       List<CTHdrFtrRef> list = sectPr.getHeaderReferenceList();
       // 回去页眉为 default的引用id
       CTHdrFtrRef df = list.stream().filter(f -> f.getType().toString().equals("default")).findFirst().get();
//       System.out.println(df.getId());

       // 获取页眉列表
       List<XWPFHeader> headers = document.getHeaderList();
       // 根据引用id,获取引用目标
//       POIXMLDocumentPart rId6 = document.getRelationById("rId6");
//       System.out.println(rId6.getParent());
       // 获取引用id
//       System.out.println(document.getRelationId(rId6));

       XWPFHeader defaultHeader = headers.stream().filter(xwpfHeader -> document.getRelationId(xwpfHeader.getPart()).equals(df.getId())).findFirst().get();
       XWPFPictureData xwpfPictureData = defaultHeader.getAllPictures().get(0);
       String fileName = xwpfPictureData.getFileName();
//       System.out.println(xwpfPictureData.toString());
//       System.out.println(fileName);
       List<XWPFParagraph> paragraphs = defaultHeader.getParagraphs();
       List<XWPFRun> runs = paragraphs.get(0).getRuns();
       for (XWPFRun r:runs
            ) {
           List<CTDrawing> drawingList = r.getCTR().getDrawingList();
           drawingList.forEach(d->d.getInlineList().forEach(l->
           {
//               System.out.println(l.getGraphic());
//               XmlObject[] xmlObjects = background.selectPath("declare namespace v='urn:schemas-microsoft-com:vml' " + ".//v:fill");
               XmlObject[] xmlObjects = l.getGraphic().selectPath("declare namespace pic='http://schemas.openxmlformats.org/drawingml/2006/picture' " + ".//pic:cNvPr");
               NamedNodeMap attributes = xmlObjects[0].getDomNode().getAttributes();
               System.out.println(attributes.getNamedItem("name").getNodeValue());

           }));
       }
//       System.out.println(defaultHeader._getHdrFtr());
       List<XWPFPictureData> allPictures = document.getAllPackagePictures();
       for (XWPFPictureData p: allPictures
            ) {
           System.out.println(p.getFileName());
       }
//       for (XWPFHeader header:headers
//            ) {
//
//           POIXMLDocumentPart part = header.getPart();
//           System.out.println(document.getRelationId(part));
//           System.out.println(header._getHdrFtr());
//           List<XWPFPictureData> allPictures = header.getAllPictures();
//           System.out.println(allPictures.get(0).getFileName());
//           XWPFPictureData xwpfPictureData = header.getAllPackagePictures().get(0);
//           System.out.println(xwpfPictureData.getFileName());
//           List<XWPFParagraph> paragraphs = header.getParagraphs();
//           List<XWPFRun> runs = paragraphs.get(0).getRuns();
//           for (XWPFRun r: runs
//                ) {
//               CTR ctr = r.getCTR();
//               List<CTDrawing> drawingList = ctr.getDrawingList();
//               CTDrawing ctDrawing = drawingList.get(0);
//               System.out.println(ctDrawing);
//           }
//       }


   }

    @Test
   public void tableTest(){
        WordUtils.setDocment("E:\\Desktop\\office\\office二级新题(34套)\\26套\\Word素材.docx");
        XWPFDocument document = WordUtils.document;
        List<XWPFTable> tables = document.getTables();
        long width = 0;
        XWPFTable xwpfTable = tables.get(0);
//        xwpfTable.getRow(0).getTableCells().stream().reduce((a,b)->{a.getWidth()});
        for (XWPFTable table : tables) {
            XWPFTableRow row = table.getRow(1);
            for (XWPFTableCell c : row.getTableCells()) {
//                System.out.println(c.getWidth());
                width += c.getWidth();
            }
        }

        System.out.println(width);
        CTDocument1 document1 = WordUtils.document.getDocument();
        CTPageSz pgSz = document1.getBody().getSectPr().getPgSz();
        System.out.println(pgSz.getW());
        long w = pgSz.getW().longValue();
        long l = w - width;
        System.out.println(l);
    }
}