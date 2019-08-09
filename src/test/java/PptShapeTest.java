import com.gzmut.office.bean.Shape;
import com.gzmut.office.enums.PowerPointConstants;
import com.gzmut.office.util.PPtUtils;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.junit.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


/**
 * 基于PPT的解析Shape测试类
 * 包含解析内容如下：
 * 获取单个幻灯片中各个Shape元素的属性如下->[List<Shape>]
 * a. id 形状元素自带ID
 * b. name 形状元素自带name
 * c. text 形状元素内文本内容
 * d. lnVal 形状元素轮廓主题
 * e. fillVal 形状元素填充主题
 * f. effectVal 形状效果主题
 * g. fontVal 形状字体主题
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-08
 */
public class PptShapeTest {
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
    private PPtUtils pptUtils = new PPtUtils(fileName);


    @Test
    public void test() throws IOException, DocumentException {
        XMLSlideShow xmlSlideShow = pptUtils.getSlideShow();
        List<Shape> shapeList;
        System.out.println("开始扫描文件：" + fileName + "中Shape元素");
        for (XSLFSlide slide : xmlSlideShow.getSlides()) {
            shapeList = pptUtils.getShape(slide);
            if (0 != shapeList.size()) {
                shapeList.forEach((shape) -> System.out.println("第" + slide.getSlideNumber() + "页幻灯片中，搜索到Shape属性：" + shape.toString()));
            }
        }
        System.out.println("扫描结束");
    }

    /**
     * 根据单个幻灯片获取其中的形状属性并包装为Shape对象
     *
     * @param slide 所需查找形状属性的幻灯片
     * @return java.util.List<bean.Video>
     */
    private List<Shape> getShape(XSLFSlide slide) throws IOException, DocumentException {
        Shape shape;
        Object attribute;
        List<Shape> list = new ArrayList<>();
        for (POIXMLDocumentPart part : slide.getSlideShow().getRelations()) {
            if (part.getPackagePart().getPartName().getName().startsWith(PowerPointConstants.SLIDE_RESOURCE_URL_PREFIX + slide.getSlideNumber())) {
                SAXReader reader = new SAXReader();
                Document document = reader.read(part.getPackagePart().getInputStream());
                Element root = document.getRootElement();
                List elements = root.selectNodes("//" + PowerPointConstants.CRUCIAL_NODE_XML_TAG);
                for (Object o : elements) {
                    shape = new Shape();
                    Element element = (Element) o;
                    attribute = element.selectSingleNode(PowerPointConstants.ID_ATTRIBUTE_KEY);
                    if (attribute != null) {
                        shape.setId(((Attribute) attribute).getValue());
                    }
                    attribute = element.selectSingleNode(PowerPointConstants.NAME_ATTRIBUTE_KEY);
                    if (attribute != null && StringUtils.isNotBlank(((Attribute) attribute).getValue())) {
                        shape.setName(((Attribute) attribute).getValue());
                    } else {
                        continue;
                    }
                    String text = getShapeElementSpliceText(element);
                    if (StringUtils.isNotBlank(text)) {
                        shape.setText(text);
                    }
                    attribute = element.selectSingleNode(PowerPointConstants.SHAPE_OUTLINE_COLOR_VALUE_XPATH);
                    if (attribute != null) {
                        shape.setLnVal(((Attribute) attribute).getValue());
                    }
                    attribute = element.selectSingleNode(PowerPointConstants.SHAPE_FILL_COLOR_VALUE_XPATH);
                    if (attribute != null) {
                        shape.setFillVal(((Attribute) attribute).getValue());
                    }
                    attribute = element.selectSingleNode(PowerPointConstants.SHAPE_EFFECT_COLOR_VALUE_XPATH);
                    if (attribute != null) {
                        shape.setEffectVal(((Attribute) attribute).getValue());
                    }
                    attribute = element.selectSingleNode(PowerPointConstants.SHAPE_FONT_COLOR_VALUE_XPATH);
                    if (attribute != null) {
                        shape.setFontVal(((Attribute) attribute).getValue());
                    }
                    if (shape.isNotBlank()) {
                        list.add(shape);
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
    private String getShapeElementSpliceText(Element shapeElement) {
        List elements = shapeElement.selectNodes("../..//a:t");
        StringBuffer stringBuffer = new StringBuffer();
        for (Object element : elements) {
            if (element instanceof Element) {
                stringBuffer.append(((Element) element).getText());
            }
        }
        return String.valueOf(stringBuffer);
    }
}