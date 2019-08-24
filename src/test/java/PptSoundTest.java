import com.gzmut.office.bean.Sound;
import com.gzmut.office.enums.PowerPointConstants;

import com.gzmut.office.util.PptUtils;
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
 * 基于PPT的解析Sound测试类
 * 包含解析内容如下：
 * 获取单个幻灯片中各个Sound元素的属性如下->[List<Sound>]
 * a. id 声音ID属性
 * b. name 声音文件名称
 * c. url 声音文件路径
 * d. showWhenStopped 声音放映时图标隐藏属性（未勾选时为NULL）
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-08
 */
public class PptSoundTest {
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
    public void test() throws IOException, DocumentException {
        List<Sound> soundList;
        XMLSlideShow xmlSlideShow = pptUtils.getSlideShow();
        System.out.println("开始扫描文件：" + fileName + "中Sound元素");
        for (XSLFSlide slide : xmlSlideShow.getSlides()) {
            soundList = pptUtils.getSound(slide);
            if (0 != soundList.size()) {
                soundList.forEach((sound) -> System.out.println("第" + slide.getSlideNumber() + "页幻灯片中，搜索到Sound属性：" + sound.toString()));
            }
        }
    }

    /**
     * 根据单个幻灯片获取其中的声音属性并包装为Sound对象
     *
     * @param slide 所需查找声音属性的幻灯片
     * @return java.util.List<bean.Sound>
     */
    private List<Sound> getSound(XSLFSlide slide) throws IOException, DocumentException {
        Sound sound;
        Object attribute;
        List<Sound> list = new ArrayList<>();
        for (POIXMLDocumentPart part : slide.getSlideShow().getRelations()) {
            if (part.getPackagePart().getPartName().getName().startsWith(PowerPointConstants.SLIDE_RESOURCE_URL_PREFIX + slide.getSlideNumber())) {
                SAXReader reader = new SAXReader();
                Document document = reader.read(part.getPackagePart().getInputStream());
                Element root = document.getRootElement();
                List elements = root.selectNodes("//" + PowerPointConstants.SLIDE_SOUND_FILE_XML_TAG);
                for (int index = 0; index < elements.size(); index++) {
                    sound = new Sound();
                    Element element = (Element) elements.get(index);
                    attribute = element.selectObject(PowerPointConstants.SLIDE_RESOURCE_LINK_ID_XML_ATTRIBUTE);
                    if (attribute != null) {
                        sound.setId(((Attribute) attribute).getValue());
                    }
                    attribute = element.selectObject("../../p:cNvPr/@name");
                    if (attribute != null) {
                        sound.setName(((Attribute) attribute).getValue());
                    }
                    attribute = element.selectSingleNode("//" + PowerPointConstants.SLIDE_SOUND_NODE_XML_TAG + "[" + index + 1 + "]/" + PowerPointConstants.SLIDE_SOUND_SHOW_WHEN_STOP_XML_ATTRIBUTE);
                    if (attribute != null) {
                        sound.setShowWhenStopped(((Attribute) attribute).getValue());
                    }
                    sound.setUrl(part.getRelationById(sound.getId()).getPackagePart().getPartName().getName());
                    if (sound.isNotBlank()) {
                        list.add(sound);
                    }
                }
            }
        }
        return list;
    }

}