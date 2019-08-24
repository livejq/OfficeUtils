package com.gzmut.office.util;

import com.gzmut.office.bean.ShapeView;
import com.gzmut.office.enums.OfficeEnums;
import com.gzmut.office.enums.PowerPointConstants;
import lombok.Getter;
import lombok.Setter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.w3c.dom.NodeList;

import java.awt.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

/**
 * MS Office 2010 pptx
 * 演示文稿的解析工具
 *
 * @author livejq
 * @since 2019/7/13
 */
@Getter
@Setter
public class PptUtils {

    /** 需要分别处理，实现原型私有 */
    private String fileName;

    private XMLSlideShow slideShow;

    public PptUtils() {
    }

    public PptUtils(String fileName) {

        this.fileName = fileName;
        if (initXMLSlideShow(fileName)) {
            System.out.println("解析工具初始化完毕！");
        } else {
            System.out.println("解析工具初始化失败~");
        }
    }

    /**
     * 获取PPt演示文档对象实例
     *
     * @return boolean
     */
    public boolean initXMLSlideShow(String fileName) {

        if (fileName == null || fileName.length() == 0) {
            return false;
        }
        if (!fileName.endsWith(OfficeEnums.PPT.getSuffix()) && !fileName.endsWith(OfficeEnums.PPTX.getSuffix())) {
            // 最好抛出异常提示信息
            return false;
        }
        this.fileName = fileName;
        // 读取PPt演示文档
        try {
            System.out.println("解析工具初始化中...");
            slideShow = new XMLSlideShow(new FileInputStream(fileName));
        } catch (IOException e) {
            System.out.println("解析工具初始化失败~");
            e.printStackTrace();
        }

        System.out.println("解析工具初始化完毕!");
        return true;
    }

    /**
     * 获取PPt演示文档对象实例
     *
     * @param fileName PPt演示文档文件路径
     */
    public void setXMLSlideShow(String fileName) {

        if (fileName == null || fileName.length() == 0) {
            return;
        }
        if (!fileName.endsWith(OfficeEnums.PPT.name()) && !fileName.endsWith(OfficeEnums.PPTX.name())) {
            // 最好抛出异常提示信息
            return;
        }
        this.fileName = fileName;
        // 读取PPt演示文档
        try {
            slideShow = new XMLSlideShow(new FileInputStream(fileName));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 获取PPt演示文档对象实例
     *
     * @return org.apache.poi.xslf.usermodel.XMLSlideShow
     */
    public XMLSlideShow getXMLSlideShow() {

        if (this.fileName == null || this.fileName.length() == 0) {
            return null;
        }
        if (!this.fileName.endsWith(OfficeEnums.PPT.name()) && !this.fileName.endsWith(OfficeEnums.PPTX.name())) {
            // 最好抛出异常提示信息
            return null;
        }
        // 读取PPt演示文档
        try {
            slideShow = new XMLSlideShow(new FileInputStream(this.fileName));
        } catch (IOException e) {
            e.printStackTrace();
        }

        return slideShow;
    }

    /**
     * 获取幻灯片总数
     *
     * @return Integer
     */
    public Integer getSlidesSum() {
        try {
            return slideShow.getSlides().size();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取PPt演示文档中的幻灯片尺寸
     *
     * @return java.awt.Dimension
     */
    public Dimension getSlideSize() {
        try {
            return slideShow.getPageSize();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取幻灯片板式
     *
     * @param slide 幻灯片对象
     * @return org.apache.poi.xslf.usermodel.SlideLayout
     */
    public String getSlideLayoutType(XSLFSlide slide) {
        try {
            return slide.getSlideLayout().getType().name();
        } catch (Exception e) {
            return "";
        }
    }

    /**
     * 获取幻灯片标题
     *
     * @param slide 幻灯片对象
     * @return java.lang.String
     */
    public String getSlideTitle(XSLFSlide slide) {
        try {
            return slide.getTitle();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取幻灯片编号
     *
     * @param slide 幻灯片对象
     * @return Integer
     */
    public Integer getSlideId(XSLFSlide slide) {
        try {
            return slide.getSlideNumber();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取幻灯片主题名称
     *
     * @param slide 幻灯片对象
     * @return java.lang.String
     */
    public String getSlideTheme(XSLFSlide slide) {
        try {
            return slide.getTheme().getName();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取幻灯片填充颜色
     *
     * @param slide 幻灯片对象
     * @return java.awt.Color
     */
    public Color getSlideBackgroundColor(XSLFSlide slide) {
        try {
            return slide.getBackground().getFillColor();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取幻灯片填充颜色的透明度
     *
     * @param slide 幻灯片对象
     * @return Integer
     */
    public Integer getSlideFillColorAlpha(XSLFSlide slide) {
        try {
            return slide.getBackground().getFillColor().getAlpha();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取某个文本框填充的颜色
     *
     * @param slide        幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @return java.awt.Color
     */
    public Color getTextBoxBackgroundColor(XSLFSlide slide, int TextBoxIndex) {
        try {
            return slide.getPlaceholder(TextBoxIndex).getFillColor();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取某个文本框中的文字内容
     *
     * @param slide        幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @return java.lang.String
     */
    public String getTextBoxContent(XSLFSlide slide, int TextBoxIndex) {
        try {
            return slide.getPlaceholder(TextBoxIndex).getText();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取某个文本框中的段落总数
     *
     * @param slide        幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @return Integer
     */
    public Integer getTextBoxParagraphSum(XSLFSlide slide, int TextBoxIndex) {
        try {
            return slide.getPlaceholder(TextBoxIndex).getTextBody().getParagraphs().size();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取某个文本框中的行高（存在单位换算问题<磅>）
     *
     * @param slide        幻灯片对象
     * @param TextBoxIndex 文本框位置编号（从0开始）
     * @return Double
     */
    public Double getTextBoxLineHeight(XSLFSlide slide, int TextBoxIndex) {
        try {
            return slide.getPlaceholder(TextBoxIndex).getTextHeight();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取某个文本框中的某一段内容
     *
     * @param slide          幻灯片对象
     * @param TextBoxIndex   文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @return java.lang.String
     */
    public String getTextBoxParagraph(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex) {
        try {
            return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getText();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取某个文本框中某一段的段前间距（存在单位换算问题<磅>）
     *
     * @param slide          幻灯片对象
     * @param TextBoxIndex   文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @return java.lang.Double
     */
    public Double getTextBoxParagraphSpaceBefore(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex) {
        try {
            return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getSpaceBefore();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取某个文本框中某一段的段后间距（存在单位换算问题<磅>）
     *
     * @param slide          幻灯片对象
     * @param TextBoxIndex   文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @return java.lang.Double
     */
    public Double getTextBoxParagraphSpaceAfter(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex) {
        try {
            return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getSpaceAfter();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取某个文本框中某一段的行距（存在单位换算问题<磅>）
     *
     * @param slide          幻灯片对象
     * @param TextBoxIndex   文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @return java.lang.Double
     */
    public Double getTextBoxParagraphLineSpace(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex) {
        try {
            return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getLineSpacing();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取某个文本框中某一段内容中的某段字符串
     *
     * @param slide          幻灯片对象
     * @param TextBoxIndex   文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @param RawIndex       字符串编号（从0开始，类型区别主要为数字、英文字母、中文等）
     * @return java.lang.String
     */
    public String getTextBoxRawContent(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex, int RawIndex) {
        try {
            return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getTextRuns().get(RawIndex).getRawText();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取某个文本框中某一段内容中的某段字符串中的超链接
     *
     * @param slide          幻灯片对象
     * @param TextBoxIndex   文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @param RawIndex       字符串编号（从0开始，类型区别主要为数字、英文字母、中文等）
     * @return java.lang.String
     */
    public String getTextBoxRawHyperlink(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex, int RawIndex) {
        try {
            return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getTextRuns().get(RawIndex).getHyperlink().getAddress();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取某个文本框中某一段内容中的某段字符串的字体大小
     *
     * @param slide          幻灯片对象
     * @param TextBoxIndex   文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @param RawIndex       字符串编号（从0开始，类型区别主要为数字、英文字母、中文等）
     * @return Float
     */
    public Float getTextBoxRawFontSize(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex, int RawIndex) {
        try {
            return Float.parseFloat(new DecimalFormat("##.#").format(slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getTextRuns().get(RawIndex).getFontSize()));
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取某个文本框中某一段内容中的某段字符串的字体样式
     *
     * @param slide          幻灯片对象
     * @param TextBoxIndex   文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @param RawIndex       字符串编号（从0开始，类型区别主要为数字、英文字母、中文等）
     * @return java.lang.String
     */
    public String getTextBoxRawFontFamily(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex, int RawIndex) {
        try {
            return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getTextRuns().get(RawIndex).getFontFamily();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取某个文本框中某一段内容中的某段字符串的字体颜色
     *
     * @param slide          幻灯片对象
     * @param TextBoxIndex   文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @param RawIndex       字符串编号（从0开始，类型区别主要为数字、英文字母、中文等）
     * @return String 返回十六进制字符串
     */
    public String getTextBoxRawFontColor(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex, int RawIndex) {
        try {
            return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getXmlObject().getRList().get(RawIndex).getRPr().getSolidFill().getSrgbClr().xgetVal().getStringValue();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 判断某个文本框中某一段内容中的某段字符串是否粗体
     *
     * @param slide          幻灯片对象
     * @param TextBoxIndex   文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @param RawIndex       字符串编号（从0开始，类型区别主要为数字、英文字母、中文等）
     * @return boolean
     */
    public boolean isTextBoxRawBold(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex, int RawIndex) {
        try {
            return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getTextRuns().get(RawIndex).isBold();
        } catch (Exception e) {
            return false;
        }
    }

    /**
     * 判断某个文本框中某一段内容中的某段字符串是否斜体
     *
     * @param slide          幻灯片对象
     * @param TextBoxIndex   文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @param RawIndex       字符串编号（从0开始，类型区别主要为数字、英文字母、中文等）
     * @return boolean
     */
    public boolean isTextBoxRawItalic(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex, int RawIndex) {
        try {
            return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getTextRuns().get(RawIndex).isItalic();
        } catch (Exception e) {
            return false;
        }
    }

    /**
     * 判断某个文本框中某一段内容中的某段字符串是否存在脚注
     *
     * @param slide          幻灯片对象
     * @param TextBoxIndex   文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @param RawIndex       字符串编号（从0开始，类型区别主要为数字、英文字母、中文等）
     * @return boolean
     */
    public boolean isTextBoxRawSubscript(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex, int RawIndex) {
        try {
            return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getTextRuns().get(RawIndex).isSubscript();
        } catch (Exception e) {
            return false;
        }
    }

    /**
     * 判断某个文本框中某一段内容中的某段字符串是否存在下划线
     *
     * @param slide          幻灯片对象
     * @param TextBoxIndex   文本框位置编号（从0开始）
     * @param ParagraphIndex 段落编号（从0开始）
     * @param RawIndex       字符串编号（从0开始，类型区别主要为数字、英文字母、中文等）
     * @return boolean
     */
    public boolean isTextBoxRawUnderlined(XSLFSlide slide, int TextBoxIndex, int ParagraphIndex, int RawIndex) {
        try {
            return slide.getPlaceholder(TextBoxIndex).getTextParagraphs().get(ParagraphIndex).getTextRuns().get(RawIndex).isUnderlined();
        } catch (Exception e) {
            return false;
        }
    }

    /**
     * 返回文本框中某个段落的字符串集合
     * @param shape 文本容器
     * @param ParagraphIndex 字符串集合下标
     * @return List<CTRegularTextRun>
     */
    public List<CTRegularTextRun> getTextRun(XSLFTextShape shape, int ParagraphIndex) {
        try {
            return shape.getTextParagraphs().get(ParagraphIndex).getXmlObject().getRList();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取pptx中的各种容器对象
     * @param pageIndex 幻灯片页码
     * @return List<XSLFShape>
     */
    public List<XSLFShape> getShape(int pageIndex) {
        try {
            return this.getSlideShow().getSlides().get(pageIndex).getShapes();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 根据单个幻灯片获取其中的形状属性并包装为Shape对象
     *
     * @param slide 所需查找形状属性的幻灯片
     * @return java.util.List<bean.Video>
     */
    public List<ShapeView> getShape(XSLFSlide slide) throws IOException, DocumentException {
        ShapeView shapeView;
        Object attribute;
        List<ShapeView> list = new ArrayList<>();
        for (POIXMLDocumentPart part : slide.getSlideShow().getRelations()) {
            if (part.getPackagePart().getPartName().getName().startsWith(PowerPointConstants.SLIDE_RESOURCE_URL_PREFIX + slide.getSlideNumber())) {
                SAXReader reader = new SAXReader();
                Document document = reader.read(part.getPackagePart().getInputStream());
                Element root = document.getRootElement();
                List elements = root.selectNodes("//" + PowerPointConstants.CRUCIAL_NODE_XML_TAG);
                for (Object o : elements) {
                    shapeView = new ShapeView();
                    Element element = (Element) o;
                    attribute = element.selectSingleNode(PowerPointConstants.ID_ATTRIBUTE_KEY);
                    if (attribute != null) {
                        shapeView.setId(((Attribute) attribute).getValue());
                    }
                    attribute = element.selectSingleNode(PowerPointConstants.NAME_ATTRIBUTE_KEY);
                    if (attribute != null && StringUtils.isNotBlank(((Attribute) attribute).getValue())) {
                        shapeView.setName(((Attribute) attribute).getValue());
                    } else {
                        continue;
                    }
                    String text = getShapeElementSpliceText(element);
                    if (StringUtils.isNotBlank(text)) {
                        shapeView.setText(text);
                    }
                    attribute = element.selectSingleNode(PowerPointConstants.SHAPE_OUTLINE_COLOR_VALUE_XPATH);
                    if (attribute != null) {
                        shapeView.setLnVal(((Attribute) attribute).getValue());
                    }
                    attribute = element.selectSingleNode(PowerPointConstants.SHAPE_FILL_COLOR_VALUE_XPATH);
                    if (attribute != null) {
                        shapeView.setFillVal(((Attribute) attribute).getValue());
                    }
                    attribute = element.selectSingleNode(PowerPointConstants.SHAPE_EFFECT_COLOR_VALUE_XPATH);
                    if (attribute != null) {
                        shapeView.setEffectVal(((Attribute) attribute).getValue());
                    }
                    attribute = element.selectSingleNode(PowerPointConstants.SHAPE_FONT_COLOR_VALUE_XPATH);
                    if (attribute != null) {
                        shapeView.setFontVal(((Attribute) attribute).getValue());
                    }
                    if (shapeView.isNotBlank()) {
                        list.add(shapeView);
                    }
                }
            }
        }
        return list;
    }

    /**
     * 根据Shape节点对象获取Shape中拼接文本信息
     *
     * @param shapeElement 所需获取文本信息的Shape Element节点对象
     * @return java.lang.String
     */
    public String getShapeElementSpliceText(Element shapeElement) {
        List elements = shapeElement.selectNodes("../..//a:t");
        StringBuffer stringBuffer = new StringBuffer();
        for (Object element : elements) {
            if (element instanceof Element) {
                stringBuffer.append(((Element) element).getText());
            }
        }
        return String.valueOf(stringBuffer);
    }

    /**
     * 返回文本框中某段落中某串字符的十六进制颜色
     * @param textRunList 某个段落的字符串集合
     * @param RawIndex 字符串集合下标
     * @return String
     */
    public String getTextRunFontColor(List<CTRegularTextRun> textRunList, int RawIndex) {
        try {
            return textRunList.get(RawIndex).getRPr().getSolidFill().getSrgbClr().xgetVal().getStringValue();
        } catch (Exception e) {
            return "";
        }
    }

    /**
     * 返回文本框中某段落中某串字符的字体样式
     * @param textRunList 某个段落的字符串集合
     * @param RawIndex 字符串集合下标
     * @return String
     */
    public String getTextRunFontStyle(List<CTRegularTextRun> textRunList, int RawIndex) {
        try {
            return textRunList.get(RawIndex).getRPr().getLatin().getTypeface();
        } catch (Exception e) {
            return "";
        }
    }

    /**
     * 返回文本框中某段落中某串字符的字体大小
     * @param textRunList 某个段落的字符串集合
     * @param RawIndex 字符串集合下标
     * @return String
     */
    public String getTextRunFontSize(List<CTRegularTextRun> textRunList, int RawIndex) {
        try {
            int fontSize = textRunList.get(RawIndex).getRPr().getSz();
            if(fontSize == 0) {
                return "";
            }
            return String.valueOf(fontSize / 100);
        } catch (Exception e) {
            return "";
        }
    }

    /**
     * 获取包含节的结点集合
     *
     * @return NodeList
     */
    public NodeList getSectionNodeList() {
        try {
            return this.getSlideShow().getCTPresentation().getDomNode().getLastChild().getFirstChild().getFirstChild().getChildNodes();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 获取每个节中所包含的幻灯片数量
     *
     * @param length 节数量
     * @return int[]
     */
    public int[] getPerSectionSlideCount(int length) {
        try {
            NodeList nodeList = this.getSectionNodeList();
            int[] temp = new int[length];
            for (int i = 0; i < length; i++) {
                temp[i] = nodeList.item(i).getFirstChild().getChildNodes().getLength();
            }
            return temp;
        } catch (Exception e) {
            return null;
        }
    }


}
