import com.gzmut.office.bean.Font;
import com.gzmut.office.bean.HyperLink;
import com.gzmut.office.enums.PowerPointConstants;

import com.gzmut.office.util.PptUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.junit.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * 基于PPT的解析SmartArt测试类
 * 包含解析内容如下：
 * 1. 获取单个幻灯片中SmartArt元素的总拼接文本内容->[String]
 * <p>
 * 2. 获取单个幻灯片中SmartArt元素的字体属性如下（Set去重）->[HashSet<Font>]
 * Set->Font Entity↓↓↓↓↓↓↓↓↓↓
 *[a. CN Font Style,
 * b. Latin Font Style,
 * c. Font Size,
 * d. Font Color,
 * e. Font Bold,
 * f. Font Tilt,
 * g. Font Underline,
 * h. Font Strike]
 * <p>
 * 3. 获取单个幻灯片中SmartArt元素的图片资源路径->[List<String>]
 * <p>
 * 4. 获取单个幻灯片中SmartArt元素的超链接属性如下->[List<HyperLink>]
 * List->HyperLinker Entity↓↓↓↓↓↓↓↓↓↓
 *[a. HyperLink resource id,
 * b. HyperLink action,
 * c. HyperLink text]
 * <p>
 * 5. 获取单个幻灯片中SmartArt元素的版式类别(type)及值(value)->[HashMap<String,String>]
 * <p>
 * 6. 获取单个幻灯片中SmartArt元素的颜色类别(type)及值(value)->[HashMap<String,String>]
 * <p>
 * 7. 获取单个幻灯片中SmartArt元素的样式类别(type)及值(value)->[HashMap<String,String>]
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-05
 */
public class PptSmartArtTest {
    /**
     * JUnit test class resource folder path
     */
    private String resourcePath = getClass().getResource("/").getPath();

    /**
     * JUnit test target PowerPoint file path
     */
    private String fileName = resourcePath + "Test.pptx";

    /**
     * PowerPoint file test tools class
     */
    private PptUtils pptUtils = new PptUtils(fileName);

    @Test
    public void test() throws IOException, XmlException, DocumentException {
        XMLSlideShow xmlSlideShow = pptUtils.getSlideShow();
        System.out.println("开始扫描文件：" + fileName + "中SmartArt元素");
        Map<String, String> smartArtAttrResult;
        HashSet<Font> smartArtFontAttrResult;
        List<String> smartArtPhotoUrlResult;
        List<HyperLink> hyperLinks;
        String smartArtTextContent;
        for (XSLFSlide slide : xmlSlideShow.getSlides()) {
            smartArtTextContent = pptUtils.getSmartArtSpliceText(slide);
            if ((StringUtils.isNotBlank(smartArtTextContent))) {
                System.out.println("第" + slide.getSlideNumber() + "页幻灯片中，搜索到SmartArt————文本内容：" + smartArtTextContent);
            }
            smartArtAttrResult = pptUtils.getSmartArtLayoutAttributes(slide, PowerPointConstants.SMART_ART_FORMAT_RESOURCE_URL_PREFIX);
            if (0 != smartArtAttrResult.size()) {
                smartArtAttrResult.forEach((type, value) -> System.out.println("第" + slide.getSlideNumber() + "页幻灯片中，搜索到SmartArt————版式类别(type)：" + type + "————版式值(value)：" + value));
            }
            smartArtAttrResult = pptUtils.getSmartArtLayoutAttributes(slide, PowerPointConstants.SMART_ART_COLOR_RESOURCE_URL_PREFIX);
            if (0 != smartArtAttrResult.size()) {
                smartArtAttrResult.forEach((type, value) -> System.out.println("第" + slide.getSlideNumber() + "页幻灯片中，搜索到SmartArt————颜色类别(type)：" + type + "———颜色值(value)：" + value));
            }
            smartArtAttrResult = pptUtils.getSmartArtLayoutAttributes(slide, PowerPointConstants.SMART_ART_STYLE_RESOURCE_URL_PREFIX);
            if (0 != smartArtAttrResult.size()) {
                smartArtAttrResult.forEach((type, value) -> System.out.println("第" + slide.getSlideNumber() + "页幻灯片中，搜索到SmartArt————样式类别(type)：" + type + "———样式值(value)：" + value));
            }
            smartArtFontAttrResult = pptUtils.getSmartArtFontStyle(slide);
            if (0 != smartArtFontAttrResult.size()) {
                smartArtFontAttrResult.forEach(str -> System.out.println("第" + slide.getSlideNumber() + "页幻灯片中，搜索到SmartArt中非默认字体属性：" + str.toString()));
            }
            smartArtPhotoUrlResult = pptUtils.getSmartArtAttributeByXpath(slide, PowerPointConstants.SMART_ART_PHOTO_XPATH);
            if (0 != smartArtPhotoUrlResult.size()) {
                smartArtPhotoUrlResult.forEach(str -> System.out.println("第" + slide.getSlideNumber() + "页幻灯片中，搜索到SmartArt————图片ID：" + str + "———图片路径：" + getSmartArtRelationshipTarget(slide, str)));
            }
            hyperLinks = pptUtils.getSmartArtHyperLink(slide);
            if (0 != hyperLinks.size()) {
                hyperLinks.forEach(str -> System.out.println("第" + slide.getSlideNumber() + "页幻灯片中，搜索到SmartArt————文本超链接：" + str.toString()));
            }
        }
        System.out.println("扫描结束");
    }

    /**
     * 根据XML Path自定义获取单个幻灯片中SmartArt元素的属性
     *
     * @param slide 所需获取属性的幻灯片对象
     * @param xPath XML Path
     * @return java.util.List<java.lang.String>
     */
    private List<String> getSmartArtAttributeByXpath(XSLFSlide slide, String xPath) throws IOException, DocumentException {
        List<String> list = new ArrayList<>();
        for (POIXMLDocumentPart part : slide.getRelations()) {
            if (part.getPackagePart().getPartName().getName().startsWith(PowerPointConstants.SMART_ART_DATA_RESOURCE_URL_PREFIX)) {
                SAXReader reader = new SAXReader();
                Document document = reader.read(part.getPackagePart().getInputStream());
                Element root = document.getRootElement();
                root.addNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                List attributes = root.selectNodes(xPath);
                for (Object attr : attributes) {
                    list.add(((Attribute) attr).getValue());
                }
            }
        }
        return list;
    }

    /**
     * 根据资源ID和幻灯片对象获取相对应的SmartArt资源Target并返回路径
     *
     * @param slide 所需获取的幻灯片对象
     * @param id    所需获取的资源ID
     * @return java.lang.String
     */
    private String getSmartArtRelationshipTarget(XSLFSlide slide, String id) {
        for (POIXMLDocumentPart part : slide.getRelations()) {
            if (part.getPackagePart().getPartName().getName().startsWith(PowerPointConstants.SMART_ART_DATA_RESOURCE_URL_PREFIX)) {
                POIXMLDocumentPart documentPart = part.getRelationById(id);
                if (documentPart != null) {
                    return documentPart.getPackagePart().getPartName().getName();
                }
            }
        }
        return null;
    }

    /**
     * 获取单个幻灯片中SmartArt元素的超链接属性并包装为HyperLink实体类对象
     *
     * @param slide 所需获取超链接属性的幻灯片对象
     * @return java.util.List<bean.HyperLink>
     */
    private List<HyperLink> getSmartArtHyperLink(XSLFSlide slide) throws IOException, DocumentException {
        List<HyperLink> links = new ArrayList<>();
        Attribute attr;
        HyperLink hyperLink;
        for (POIXMLDocumentPart part : slide.getRelations()) {
            if (part.getPackagePart().getPartName().getName().startsWith(PowerPointConstants.SMART_ART_DATA_RESOURCE_URL_PREFIX)) {
                SAXReader reader = new SAXReader();
                Document document = reader.read(part.getPackagePart().getInputStream());
                Element root = document.getRootElement();
                // 根据"/"路径获取元素
                List elements = root.selectNodes("//" + PowerPointConstants.SMART_ART_HYPERLINK_XML_TAG);
                for (Object element : elements) {
                    hyperLink = new HyperLink();
                    attr = ((Attribute) ((Element) element).selectSingleNode(PowerPointConstants.SMART_ART_HYPERLINK_RID_NAME));
                    if (StringUtils.isNotBlank(String.valueOf(attr.getValue()))) {
                        hyperLink.setId(attr.getValue());
                    }
                    attr = ((Attribute) ((Element) element).selectSingleNode(PowerPointConstants.SMART_ART_HYPERLINK_ACTION_NAME));
                    if (attr != null) {
                        hyperLink.setLinkAction(attr.getValue());
                    }
                    Iterator iterator = ((Element) element).getParent().getParent().elementIterator();
                    if (iterator.hasNext()) {
                        iterator.next();
                        hyperLink.setLinkText(((Element) iterator.next()).getText());
                    }
                    if (hyperLink.isNotBlank()) {
                        links.add(hyperLink);
                    }
                }
            }
        }
        return links;
    }

    /**
     * 获取单个幻灯片中SmartArt元素的字体属性如下（Set去重）
     *
     * @param slide 所需获取字体属性的幻灯片
     * @return java.util.HashSet<bean.Font>
     */
    private HashSet<Font> getSmartArtFontStyle(XSLFSlide slide) throws IOException, DocumentException {
        HashSet<Font> set = new HashSet<>();
        Font font;
        Attribute attr;
        for (POIXMLDocumentPart part : slide.getRelations()) {
            if (part.getPackagePart().getPartName().getName().startsWith(PowerPointConstants.SMART_ART_DATA_RESOURCE_URL_PREFIX)) {
                SAXReader reader = new SAXReader();
                Document document = reader.read(part.getPackagePart().getInputStream());
                Element root = document.getRootElement();
                // 根据"/"路径获取元素
                List elements = root.selectNodes("//" + PowerPointConstants.SMART_ART_TEXT_NODE_TAG);
                for (Object element : elements) {
                    font = new Font();
                    attr = ((Attribute) ((Element) element).selectSingleNode(PowerPointConstants.SMART_ART_TEXT_EA_STYLE_XPATH));
                    if (attr != null) {
                        font.setEaStyle(attr.getValue());
                    }
                    attr = ((Attribute) ((Element) element).selectSingleNode(PowerPointConstants.SMART_ART_TEXT_LATIN_STYLE_XPATH));
                    if (attr != null) {
                        font.setLatinStyle(attr.getValue());
                    }
                    attr = ((Attribute) ((Element) element).selectSingleNode(PowerPointConstants.SMART_ART_TEXT_SIZE_NAME));
                    if (attr != null) {
                        font.setSize(attr.getValue());
                    }
                    attr = ((Attribute) ((Element) element).selectSingleNode(PowerPointConstants.SMART_ART_TEXT_COLOR_XPATH));
                    if (attr != null) {
                        font.setColor(attr.getValue());
                    }
                    attr = ((Attribute) ((Element) element).selectSingleNode(PowerPointConstants.SMART_ART_TEXT_BOLD_NAME));
                    if (attr != null) {
                        font.setBold(attr.getValue());
                    }
                    attr = ((Attribute) ((Element) element).selectSingleNode(PowerPointConstants.SMART_ART_TEXT_TILT_NAME));
                    if (attr != null) {
                        font.setTilt(attr.getValue());
                    }
                    attr = ((Attribute) ((Element) element).selectSingleNode(PowerPointConstants.SMART_ART_TEXT_UNDERLINE_NAME));
                    if (attr != null) {
                        font.setUnderline(attr.getValue());
                    }
                    attr = ((Attribute) ((Element) element).selectSingleNode(PowerPointConstants.SMART_ART_TEXT_STRIKE_NAME));
                    if (attr != null) {
                        font.setStrike(attr.getValue());
                    }
                    if (font.isNotBlank()) {
                        set.add(font);
                    }
                }
            }
        }
        return set;
    }

    /**
     * 获取单个幻灯片中SmartArt元素的总拼接文本内容
     *
     * @param slide 所需获取文本内容的幻灯片对象
     * @return java.lang.String
     */
    private String getSmartArtSpliceText(XSLFSlide slide) throws IOException, XmlException {
        StringBuffer stringBuffer = new StringBuffer();
        XmlObject xmlObject;
        XmlCursor xmlCursor;
        for (POIXMLDocumentPart part : slide.getRelations()) {
            if (part.getPackagePart().getPartName().getName().startsWith(PowerPointConstants.SMART_ART_DATA_RESOURCE_URL_PREFIX)) {
                xmlObject = XmlObject.Factory.parse(part.getPackagePart().getInputStream());
                xmlCursor = xmlObject.newCursor();
                while (xmlCursor.hasNextToken()) {
                    if (xmlCursor.isText()) {
                        stringBuffer.append(xmlCursor.getTextValue());
                    }
                    xmlCursor.toNextToken();
                }
            }
        }
        return String.valueOf(stringBuffer);
    }

    /**
     * 根据所传递的资源路径获取单个幻灯片中SmartArt元素的对应属性，
     * 以[type:pri]键值(value)形式存在取出包装并返回
     *
     * @param slide       所需获取属性的幻灯片对象
     * @param resourceUrl 所需获取属性的资源文件路径，例：../slide1.xml,../data1.xml,../color.xml
     * @return java.util.Map<java.lang.String, java.lang.String>
     */
    private Map<String, String> getSmartArtLayoutAttributes(XSLFSlide slide, String resourceUrl) throws IOException, XmlException {
        XmlObject xmlObject;
        XmlCursor xmlCursor;
        String key;
        Map<String, String> map = new HashMap<>();
        for (POIXMLDocumentPart part : slide.getRelations()) {
            if (part.getPackagePart().getPartName().getName().startsWith(resourceUrl)) {
                xmlObject = XmlObject.Factory.parse(part.getPackagePart().getInputStream());
                xmlCursor = xmlObject.newCursor();
                while (xmlCursor.hasNextToken()) {
                    if (PowerPointConstants.SMART_ART_TYPE_KEY.equals(String.valueOf(xmlCursor.getName()))) {
                        key = xmlCursor.getTextValue();
                        xmlCursor.toNextToken();
                        if (PowerPointConstants.SMART_ART_PRIMARY_KEY.equals(String.valueOf(xmlCursor.getName()))) {
                            map.put(key, xmlCursor.getTextValue());
                            if (!xmlCursor.hasNextToken()) {
                                break;
                            }
                        }
                    }
                    xmlCursor.toNextToken();
                }
            }
        }
        return map;
    }
}
