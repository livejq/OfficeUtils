import com.gzmut.office.bean.Video;
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
 * 基于PPT的解析Video测试类
 * 包含解析内容如下：
 * 获取单个幻灯片中各个Video元素的属性如下->[List<Video>]
 * a. id 视频ID属性
 * b. name 视频文件名称
 * c. url 视频文件路径
 * d. coverId 视频封面图片ID属性
 * e. coverUrl 视频封面文件路径
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-08
 */
public class PptVideoTest {
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
        List<Video> videoList;
        XMLSlideShow xmlSlideShow = pptUtils.getSlideShow();
        System.out.println("开始扫描文件：" + fileName + "中Video元素");
        for (XSLFSlide slide : xmlSlideShow.getSlides()) {
            videoList = pptUtils.getVideo(slide);
            if (0 != videoList.size()) {
                videoList.forEach((video) -> System.out.println("第" + slide.getSlideNumber() + "页幻灯片中，搜索到Video属性：" + video.toString()));
            }
        }
        System.out.println("扫描结束");
    }

    /**
     * 根据单个幻灯片获取其中的视频属性并包装为Video对象
     *
     * @param slide 所需查找视频属性的幻灯片
     * @return java.util.List<bean.Video>
     */
    private List<Video> getVideo(XSLFSlide slide) throws IOException, DocumentException {
        Video video;
        Object attribute;
        List<Video> list = new ArrayList<>();
        for (POIXMLDocumentPart part : slide.getSlideShow().getRelations()) {
            if (part.getPackagePart().getPartName().getName().startsWith(PowerPointConstants.SLIDE_RESOURCE_URL_PREFIX + slide.getSlideNumber())) {
                SAXReader reader = new SAXReader();
                Document document = reader.read(part.getPackagePart().getInputStream());
                Element root = document.getRootElement();
                List elements = root.selectNodes("//" + PowerPointConstants.SLIDE_VIDEO_XML_TAG);
                for (Object o : elements) {
                    video = new Video();
                    Element element = (Element) o;
                    attribute = element.selectObject(PowerPointConstants.SLIDE_RESOURCE_LINK_ID_XML_ATTRIBUTE);
                    if (attribute != null) {
                        video.setId(((Attribute) attribute).getValue());
                    }
                    attribute = element.selectObject("../../p:cNvPr/@name");
                    if (attribute != null) {
                        video.setName(((Attribute) attribute).getValue());
                    }
                    attribute = element.selectObject("../../..//a:blip/@r:embed");
                    if (attribute != null) {
                        video.setCoverId(((Attribute) attribute).getValue());
                    }
                    video.setUrl(part.getRelationById(video.getId()).getPackagePart().getPartName().getName());
                    video.setCoverUrl(part.getRelationById(video.getCoverId()).getPackagePart().getPartName().getName());
                    if (video.isNotBlank()) {
                        list.add(video);
                    }
                }
            }
        }
        return list;
    }
}