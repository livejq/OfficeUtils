package com.gzmut.office.util;

import com.gzmut.office.enums.ppt.PptSectionPropertiesEnums;
import com.gzmut.office.enums.ppt.PptSlidePropertiesEnums;
import com.gzmut.office.enums.ppt.PptTargetEnums;
import com.gzmut.office.enums.ppt.PptTextBoxPropertiesEnums;
import com.gzmut.office.bean.Location;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xslf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.w3c.dom.NodeList;

import java.awt.*;
import java.util.*;
import java.util.List;

/**
 * 根据知识点对pptx
 * 演示文稿进行判题
 *
 * @author livejq
 * @since 2019/7/31
 */
@Slf4j
@Getter
@Setter
public class PptCheckUtils {

    private PptUtils pptUtils;

    public PptCheckUtils() {
    }

    public PptCheckUtils(PptUtils pptUtils) {
        this.pptUtils = pptUtils;
    }

    public boolean initPptUtils(String fileName) {
        return this.pptUtils.initXMLSlideShow(fileName);
    }

    /**
     * 判断（幻灯片/文本框等）背景颜色是否正确
     *
     * @param location 定位参数
     * @param param    hex颜色
     * @param score    小题总分
     * @param sumRule  小题规则总数
     * @return Map
     */
    public Map<String, String> checkSlideBackgroundColor(Location location, Map<String, Object> param, String score, int sumRule) {

        Map<String, String> infoMap = new HashMap<>(2);
        String[] slides = location.getLp().split(",");

        if (slides != null) {
            for (String slide : slides) {
                PptTargetEnums target = PptTargetEnums.valueOf(location.getLo());
                switch (target) {
                    case SLIDE:
                        int red, green, blue;
                        Color color = pptUtils.getSlideBackgroundColor(pptUtils.getSlideShow()
                                .getSlides()
                                .get(Integer.parseInt(slide) - 1));
                        if (color != null) {
                            red = color.getRed();
                            green = color.getGreen();
                            blue = color.getBlue();
                        } else {
                            red = green = blue = -1;
                        }

                        String backgroundColor = param.get("BACKGROUND_COLOR").toString();
                        String[] answer_color = backgroundColor.split(",");

                        if (red == Integer.parseInt(answer_color[0]) && green == Integer.parseInt(answer_color[1]) && blue == Integer.parseInt(answer_color[2])) {
                            infoMap.put("score", String.valueOf(Integer.parseInt(score) / sumRule));
                            infoMap.put("msg", "背景颜色正确!");
                        } else {
                            infoMap.put("score", "0.0");
                            infoMap.put("msg", "背景颜色错误!");
                        }

                        return infoMap;
                    case TEXT_BOX:
                        return null;

                    default:
                        return null;
                }
            }
        }

        return null;
    }

    /**
     * 检查幻灯片数量
     *
     * @param location 定位参数
     * @param param    元素数量
     * @param score    小题总分
     * @param sumRule  小题规则总数
     * @return Map
     */
    public Map<String, String> checkSlideNumber(Location location, Map<String, Object> param, String score, int sumRule) {

        Map<String, String> infoMap = new HashMap<>(2);
        int answer_sumPage = Integer.parseInt(param.get("COUNT").toString());
        int sumPage = pptUtils.getSlidesSum();

        if (sumPage == answer_sumPage) {
            Float temp = Float.parseFloat(score) / sumRule;
            infoMap.put("score", String.valueOf(temp));
            infoMap.put("msg", "幻灯片数量正确：" + answer_sumPage + "张;" + System.lineSeparator() + System.lineSeparator());
        } else {
            infoMap.put("score", "0.0");
            infoMap.put("msg", "幻灯片数量错误!" + "需要：" + answer_sumPage + "张," + "实际：" + sumPage + "张;" + System.lineSeparator() + System.lineSeparator());
        }

        return infoMap;
    }

    /**
     *
     * @param location 定位参数
     * @param param    元素数量
     * @param score    小题总分
     * @param sumRule  小题规则总数
     * @return Map
     */
    public Map<String, String> checkSwitchStyle(Location location, Map<String, Object> param, String score, int sumRule) {

        Map<String, String> infoMap = new HashMap<>(2);

        return infoMap;
    }

    /**
     * 检查文本超链接
     *
     * @param location 定位参数
     * @param param    元素数量
     * @param score    小题总分
     * @param sumRule  小题规则总数
     * @return Map
     */
    public Map<String, String> checkTextHyperlink(Location location, Map<String, Object> param, String score, int sumRule) {

        Map<String, String> infoMap = new HashMap<>(2);
        int answer_textLink = Integer.parseInt(param.get("TEXT_HYPERLINK").toString());
        int num = 0;
        StringBuilder sb = new StringBuilder("[文本超链接检查结果] " + System.lineSeparator());
        List<XSLFSlide> slides = pptUtils.getSlideShow().getSlides();
        if (slides.size() != 0) {
            int page = Integer.parseInt(location.getLp());
            sb.append("第 " + page + " 张幻灯片:");
            List<XSLFShape> xslfShapeList = pptUtils.getShape(page - 1);
            for (int j = 0, len = xslfShapeList.size(); j < len; j++) {
                XSLFShape sh = xslfShapeList.get(j);
                if (sh instanceof XSLFTextShape) {
                    XSLFTextShape shape = (XSLFTextShape) sh;
                    String xml = shape.getXmlObject().xmlText();
                    int index = 0;
                    String child = "hlinksldjump";
                    while ((index = xml.indexOf(child, index)) != -1) {
                        index = index + child.length();
                        num++;
                    }
                }
            }
            if (num == answer_textLink) {
                // 此规则的总分数
                sb.append("文本超链接设置正确");
                float temp = Float.parseFloat(score) / sumRule;
                infoMap.put("score", String.valueOf(temp));
                infoMap.put("msg", sb.toString());
            } else {
                sb.append("文本超链接设置错误");
                infoMap.put("score", "0.0");
                infoMap.put("msg", sb.toString());
            }
        }

        return infoMap;
    }

    /**
     * 检查主题
     *
     * @param location 定位参数
     * @param param    元素数量
     * @param score    小题总分
     * @param sumRule  小题规则总数
     * @return Map
     */
    public Map<String, String> checkSlideTheme(Location location, Map<String, Object> param, String score, int sumRule) {

        Map<String, String> infoMap = new HashMap<>(2);
        String[] pages = location.getLp().split(",");
        String answer_theme = param.get("THEME").toString();
        int num = 0;
        StringBuilder sb = new StringBuilder("[幻灯片主题检查结果] " + System.lineSeparator());
        List<XSLFSlide> slides = pptUtils.getSlideShow().getSlides();
        if (slides.size() != 0) {
            if (pages.length == 1) {
                int page = Integer.parseInt(pages[0]);
                sb.append("第 " + page + " 张幻灯片:");
                String theme = pptUtils.getSlideTheme(slides.get(page - 1));
                if (theme.equals(answer_theme)) {
                    sb.append("主题设置正确" + System.lineSeparator());
                    // 此规则的总分数
                    float temp = Float.parseFloat(score) / sumRule;
                    infoMap.put("score", String.valueOf(temp));
                    infoMap.put("msg", sb.toString());
                } else {
                    sb.append("主题设置错误" + System.lineSeparator());
                    infoMap.put("score", "0.0");
                    infoMap.put("msg", sb.toString());
                }

                return infoMap;
            } else {
                int start_page = Integer.parseInt(pages[0]);
                int end_page = Integer.parseInt(pages[1]);
                int index = start_page;
                while (index <= end_page) {
                    sb.append("第 " + index + " 张幻灯片:");
                    String theme = pptUtils.getSlideTheme(slides.get(index - 1));
                    if (theme.equals(answer_theme)) {
                        num++;
                    }
                    index++;
                }
                if (num == (end_page - start_page + 1)) {
                    sb.append("主题设置正确" + System.lineSeparator());
                    // 此规则的总分数
                    float temp = Float.parseFloat(score) / sumRule;
                    infoMap.put("score", String.valueOf(temp));
                    infoMap.put("msg", sb.toString());
                } else {
                    sb.append("主题设置错误" + System.lineSeparator());
                    infoMap.put("score", "0.0");
                    infoMap.put("msg", sb.toString());
                }

                return infoMap;
            }
        } else {
            return null;
        }
    }

    /**
     * 检查幻灯片版式风格
     *
     * @param location 定位参数
     * @param param    元素数量
     * @param score    小题总分
     * @param sumRule  小题规则总数
     * @return Map
     */
    public Map<String, String> checkLayoutStyle(Location location, Map<String, Object> param, String score, int sumRule) {
        Map<String, String> infoMap = new HashMap<>(2);
        String slide = location.getLp();
        int num = 0;
        StringBuilder sb = new StringBuilder("[幻灯片板式风格检查结果]");
        if (slide != null || slide.length() != 0) {
            String answer_layout = param.get(PptSlidePropertiesEnums.LAYOUT_STYLE.getProperty()).toString();
            String layout = pptUtils.getSlideLayoutType(pptUtils.getSlideShow().getSlides()
                    .get(Integer.parseInt(slide) - 1));
            if (answer_layout != null || answer_layout.length() != 0) {
                if (answer_layout.equals(layout)) {
                    num++;
                    sb.append("演示文稿中第 " + slide + " 张幻灯片版式正确," + answer_layout + System.lineSeparator() + System.lineSeparator());
                }
            }
            // 此规则的总分数
            float temp = Float.parseFloat(score) / sumRule;
            infoMap.put("score", String.valueOf(temp * num));
            infoMap.put("msg", sb.toString());

            return infoMap;
        }

        return null;
    }

    /**
     * 检查节名称和数量
     *
     * @param location 定位参数
     * @param param    元素数量
     * @param score    小题总分
     * @param sumRule  小题规则总数
     * @return Map
     */
    public Map<String, String> checkSection(Location location, Map<String, Object> param, String score, int sumRule) {

        Map<String, String> infoMap = new HashMap<>(2);
        String[] answer_section_name = param.get(PptSectionPropertiesEnums.SECTION_NAME.getProperty()).toString().split("&#@@#&");
        String[] answer_section_number = param.get(PptSectionPropertiesEnums.SECTION_NUMBER.getProperty()).toString().split(",");
        String xml = pptUtils.getSlideShow().getCTPresentation().xmlText();
        StringBuilder sb = new StringBuilder("[ppt节检查结果]");
        boolean result = true;
        for (String name : answer_section_name) {
            if (!xml.contains(name)) {
                result = false;
                break;
            }
        }
        int len = answer_section_number.length;
        NodeList nodeList = pptUtils.getSectionNodeList();
        if (nodeList != null) {
            if (nodeList.getLength() != len) {
                result = false;
            } else {
                for (int i = 0; i < len; i++) {
                    if (nodeList.item(i).getFirstChild().getChildNodes().getLength() != Integer.parseInt(answer_section_number[i])) {
                        result = false;
                        break;
                    }
                }
            }
        }
        if (result) {
            sb.append("节设置正确");
            // 此规则的总分数
            float temp = Float.parseFloat(score) / sumRule;
            infoMap.put("score", String.valueOf(temp));
            infoMap.put("msg", sb.toString());
        } else {
            sb.append("节设置错误");
            infoMap.put("score", "0.0");
            infoMap.put("msg", sb.toString());
        }

        return infoMap;
    }

    /**
     * 检查幻灯片母版名称
     *
     * @param location 定位参数
     * @param param    元素数量
     * @param score    小题总分
     * @param sumRule  小题规则总数
     * @return Map
     */
    public Map<String, String> checkLayoutName(Location location, Map<String, Object> param, String score, int sumRule) {
        Map<String, String> infoMap = new HashMap<>(2);
        List<String> layout_name = new LinkedList<>();
        int num = 0;
        StringBuilder sb = new StringBuilder("[幻灯片板式名称检查结果]");
        String answer_layout_name = param.get(PptSlidePropertiesEnums.LAYOUT_NAME.getProperty()).toString();
        List<XSLFSlideMaster> masters = pptUtils.getSlideShow().getSlideMasters();
        for (XSLFSlideMaster master : masters) {
            for (XSLFSlideLayout layout : master.getSlideLayouts()) {
                layout_name.add(layout.getType().toString());
            }
        }
        for (String name : layout_name) {
            if (!answer_layout_name.contains(name)) {
                num++;
                sb.append("错误" + System.lineSeparator() + System.lineSeparator());
                break;
            }
        }

        // 此规则的总分数
        float temp = Float.parseFloat(score) / sumRule;
        if (num == 0) {
            sb.append("正确" + System.lineSeparator() + System.lineSeparator());
            infoMap.put("score", String.valueOf(temp));
            infoMap.put("msg", sb.toString());
        } else {
            infoMap.put("score", "0.0");
            infoMap.put("msg", sb.toString());
        }

        return infoMap;
    }

    /**
     * 检查目标元素的数量
     *
     * @param location 定位参数
     * @param param    元素数量
     * @param score    小题总分
     * @param sumRule  小题规则总数
     * @return Map
     */
    public Map<String, String> checkTextBoxContent(Location location, Map<String, Object> param, String score, int sumRule) {

        Map<String, String> infoMap = new HashMap<>(2);
        StringBuilder sb = new StringBuilder("[ppt文本内容检查结果]" + System.lineSeparator());
        int num = 0;
        int textBoxCount = 0;
        int page = Integer.parseInt(location.getLp());
        String[] answer_content = param.get(PptTextBoxPropertiesEnums
                .TEXT_CONTENT.getProperty()).toString().split("&#@@#&");

        sb.append("第 " + page + " 张幻灯片:");
        List<XSLFShape> xslfShapeList = pptUtils.getShape(page - 1);
        if (xslfShapeList != null && xslfShapeList.size() != 0) {
            for (int j = 0, len = xslfShapeList.size(); j < len; j++) {
                XSLFShape sh = xslfShapeList.get(j);
                for (String answer : answer_content) {
                    if (sh instanceof XSLFTextShape) {
                        textBoxCount++;
                        XSLFTextShape shape = (XSLFTextShape) sh;
                        if (shape.getText().contains(answer)) {
                            sb.append("文本内容正确:" + answer + System.lineSeparator() + System.lineSeparator());
                            num++;
                        }
                    }
                }
            }
        }
        // 此规则的总分数
        float temp = Float.parseFloat(score) / sumRule;
        String finalScore = String.valueOf(temp * num / textBoxCount);
        infoMap.put("score", finalScore);
        infoMap.put("msg", sb.toString());

        return infoMap;
    }

    public Map<String, String> checkTextAlignStyle(Location location, Map<String, Object> param, String score, int sumRule) {

        Map<String, String> infoMap = new HashMap<>(2);
        StringBuilder sb = new StringBuilder("[文本框对齐方式检查结果]" + System.lineSeparator());
        int[] h_align = new int[3];
        int[] v_align = new int[3];
        boolean result = true;
        int page = Integer.parseInt(location.getLp());
        // 三个参数，左中右个数
        String[] answer_h_align = param.get(PptTextBoxPropertiesEnums
                .HORIZONTAL_ALIGN.getProperty()).toString().split(",");
        // 三个参数，上中下个数
        String[] answer_v_align = param.get(PptTextBoxPropertiesEnums
                .VERTICAL_ALIGN.getProperty()).toString().split(",");

        sb.append("第 " + page + " 张幻灯片:");
        List<XSLFShape> xslfShapeList = pptUtils.getShape(page - 1);
        if (xslfShapeList != null && xslfShapeList.size() != 0) {
            for (int j = 0, len = xslfShapeList.size(); j < len; j++) {
                XSLFShape sh = xslfShapeList.get(j);
                if (sh instanceof XSLFTextShape) {
                    XSLFTextShape shape = (XSLFTextShape) sh;
                    if (shape.getText().length() != 0) {
                        String xml = shape.getXmlObject().xmlText();
                        int index = 0;
                        int num = 0;
                        // 水平判断
                        String child = "algn=\"l\"";
                        while ((index = xml.indexOf(child, index)) != -1) {
                            index = index + child.length();
                            num++;
                        }
                        h_align[0] += num;

                        index = 0;
                        num = 0;
                        child = "algn=\"ctr\"";
                        while ((index = xml.indexOf(child, index)) != -1) {
                            index = index + child.length();
                            num++;
                        }
                        h_align[1] += num;

                        index = 0;
                        num = 0;
                        child = "algn=\"r\"";
                        while ((index = xml.indexOf(child, index)) != -1) {
                            index = index + child.length();
                            num++;
                        }
                        h_align[2] += num;

                        // 垂直判断
                        index = 0;
                        num = 0;
                        child = "anchor=\"l\"";
                        while ((index = xml.indexOf(child, index)) != -1) {
                            index = index + child.length();
                            num++;
                        }
                        v_align[0] += num;

                        index = 0;
                        num = 0;
                        child = "anchor=\"ctr\"";
                        while ((index = xml.indexOf(child, index)) != -1) {
                            index = index + child.length();
                            num++;
                        }
                        v_align[1] += num;

                        index = 0;
                        num = 0;
                        child = "anchor=\"r\"";
                        while ((index = xml.indexOf(child, index)) != -1) {
                            index = index + child.length();
                            num++;
                        }
                        v_align[2] += num;
                    }
                }
            }
        }
        for (int i = 0, len = 3; i < len; i++) {
            if (h_align[i] != Integer.parseInt(answer_h_align[i])) {
                result = false;
                break;
            }
            if (v_align[i] != Integer.parseInt(answer_v_align[i])) {
                result = false;
                break;
            }
        }
        if (result) {
            sb.append("文本段落对齐正确");
            // 此规则的总分数
            float temp = Float.parseFloat(score) / sumRule;
            infoMap.put("score", String.valueOf(temp));
            infoMap.put("msg", sb.toString());
        } else {
            sb.append("文本段落对齐错误");
            // 此规则的总分数
            infoMap.put("score", "0.0");
            infoMap.put("msg", sb.toString());
        }

        return infoMap;
    }

    /**
     * 检查文本框中的文本格式
     *
     * @param location 定位参数
     * @param param    元素数量
     * @param score    小题总分
     * @param sumRule  小题规则总数
     * @return Map
     */
    public Map<String, String> checkTextBoxContentFormat(Location location, Map<String, Object> param, String score, int sumRule) {

        Map<String, String> infoMap = new HashMap<>(2);
        // slides共三个值：[0] 开始页，[1]中止页
        String[] pages = location.getLp().split(",");
        int num = 0;
        int textBoxCount = 0;
        StringBuilder sb = new StringBuilder("[文本格式检查结果] " + System.lineSeparator());
        List<XSLFSlide> slides = pptUtils.getSlideShow().getSlides();
        Random random = new Random();
        if (slides.size() != 0) {
            if (pages.length == 1) {
                int page = Integer.parseInt(pages[0]);
                sb.append("第 " + page + " 张幻灯片:");
                List<XSLFShape> xslfShapeList = pptUtils.getShape(page - 1);
                for (int j = 0, len = xslfShapeList.size(); j < len; j++) {
                    XSLFShape sh = xslfShapeList.get(j);
                    if (sh instanceof XSLFTextShape) {
                        textBoxCount++;
                        XSLFTextShape shape = (XSLFTextShape) sh;
                        if (shape.getTextParagraphs().size() == 0) {
                            return null;
                        }
                        int randPara = random.nextInt(shape.getTextParagraphs().size());
                        List<CTRegularTextRun> rList = shape.getTextParagraphs().get(randPara).getXmlObject().getRList();
                        int randRaw = rList.size() == 0 ? 0 : random.nextInt(rList.size());
                        // 判断字体颜色
                        String hexColor = pptUtils.getTextRunFontColor(rList, randRaw);
                        Object strColor = param.get(PptTextBoxPropertiesEnums
                                .TEXT_COLOR.getProperty());
                        String answer_hexColor = (strColor == null ? "" : strColor.toString());
                        if (hexColor.contains(answer_hexColor)) {
                            num++;
                            if (answer_hexColor.length() == 0) {
                                answer_hexColor = "默认值";
                            }
                            sb.append("文本框中字体颜色正确：" + answer_hexColor + System.lineSeparator());
                        }
                        // 判断字体样式
                        String fontStyle = pptUtils.getTextRunFontStyle(rList, randRaw);
                        Object strStyle = param.get(PptTextBoxPropertiesEnums
                                .TEXT_STYLE.getProperty());
                        String answer_fontStyle = (strStyle == null ? "" : strStyle.toString());
                        if (answer_fontStyle.equals(fontStyle)) {
                            num++;
                            if (answer_fontStyle.length() == 0) {
                                answer_fontStyle = "默认值";
                            }
                            sb.append("文本框中字体样式正确:" + answer_fontStyle + System.lineSeparator());
                        }
                        // 判断字体大小
                        String fontSize = pptUtils.getTextRunFontSize(rList, randRaw);
                        Object strSize = param.get(PptTextBoxPropertiesEnums
                                .TEXT_SIZE.getProperty());
                        String answer_fontSize = strSize.toString();
                        if (answer_fontSize.equals(fontSize)) {
                            num++;
                            if (answer_fontSize.length() == 0) {
                                answer_fontSize = "默认值";
                            }
                            sb.append("文本框中字体大小正确:" + answer_fontSize + System.lineSeparator() + System.lineSeparator());
                        }
                        // 判断字体粗体
                        // 判断字体斜体
                    }
                }
            } else {
                int start_page = Integer.parseInt(pages[0]);
                int end_page = Integer.parseInt(pages[1]);
                int index = start_page;
                while (index <= end_page) {
                    sb.append("第 " + index + " 张幻灯片:");
                    List<XSLFShape> xslfShapeList = pptUtils.getShape(index - 1);
                    index++;
                    for (int j = 0, len = xslfShapeList.size(); j < len; j++) {
                        XSLFShape sh = xslfShapeList.get(j);
                        if (sh instanceof XSLFTextShape) {
                            textBoxCount++;
                            XSLFTextShape shape = (XSLFTextShape) sh;
                            if (shape.getTextParagraphs().size() == 0) {
                                return null;
                            }
                            int randPara = random.nextInt(shape.getTextParagraphs().size());
                            List<CTRegularTextRun> rList = pptUtils.getTextRun(shape, randPara);
                            int randRaw = rList.size() == 0 ? 0 : random.nextInt(rList.size());
                            // 判断字体颜色
                            String hexColor = pptUtils.getTextRunFontColor(rList, randRaw);
                            String strColor = param.get(PptTextBoxPropertiesEnums.TEXT_COLOR.getProperty()) == null ? ""
                                    : param.get(PptTextBoxPropertiesEnums.TEXT_COLOR.getProperty()).toString();
                            String answer_hexColor = strColor;
                            if (hexColor.contains(answer_hexColor)) {
                                num++;
                                if (answer_hexColor.length() == 0) {
                                    answer_hexColor = "默认值";
                                }
                                sb.append("文本框中字体颜色正确:" + answer_hexColor + System.lineSeparator());
                            }
                            // 判断字体样式
                            String fontStyle = pptUtils.getTextRunFontStyle(rList, randRaw);
                            Object strStyle = param.get(PptTextBoxPropertiesEnums
                                    .TEXT_STYLE.getProperty());
                            String answer_fontStyle = (strStyle == null ? "" : strStyle.toString());
                            if (answer_fontStyle.equals(fontStyle)) {
                                num++;
                                if (answer_fontStyle.length() == 0) {
                                    answer_fontStyle = "默认值";
                                }
                                sb.append("文本框中字体样式正确:" + answer_fontStyle + System.lineSeparator());
                            }
                            // 判断字体大小
                            String fontSize = pptUtils.getTextRunFontSize(rList, randRaw);

                            Object strSize = param.get(PptTextBoxPropertiesEnums
                                    .TEXT_SIZE.getProperty());
                            String answer_fontSize = strSize.toString();
                            if (answer_fontSize.equals(fontSize)) {
                                num++;
                                if (answer_fontSize.length() == 0) {
                                    answer_fontSize = "默认值";
                                }
                                sb.append("文本框中字体大小正确:" + answer_fontSize + System.lineSeparator() + System.lineSeparator());
                            }

                            // 判断字体粗体
                            // 判断字体斜体
                        }

                    }
                }
            }
        }
        // 每个文本框检查3个属性
        float correct = (float) (num / textBoxCount * 3);
        // 此规则的总分数
        float temp = Float.parseFloat(score) / sumRule;
        String finalScore = String.valueOf(temp * correct);
        infoMap.put("score", finalScore);
        infoMap.put("msg", sb.toString());

        return infoMap;
    }

    /**
     * 检查文本框中的文本格式
     *
     * @param location 定位参数
     * @param param    元素数量
     * @param score    小题总分
     * @param sumRule  小题规则总数
     * @return Map
     */
    public Map<String, String> checkTextLevel(Location location, Map<String, Object> param, String score, int sumRule) {

        Map<String, String> infoMap = new HashMap<>(2);
        int page = 0;
        if (location.getLp() != null) {
            page = Integer.parseInt(location.getLp());
        }
        if (page == 0) {
            return null;
        } else {
            StringBuilder sb = new StringBuilder("[文本级别检查结果] " + System.lineSeparator());
            List<XSLFSlide> slides = pptUtils.getSlideShow().getSlides();
            sb.append("第 " + page + " 张幻灯片:");
            List<XSLFShape> xslfShapeList = pptUtils.getShape(page - 1);
            if (xslfShapeList.size() != 0) {
                int title = 0;
                int subTitle = 0;
                int first_text = 0;
                int second_text = 0;
                int third_text = 0;
                int answer_count = param.size();
                int count = 0;
                for (int j = 0, len = xslfShapeList.size(); j < len; j++) {
                    XSLFShape sh = xslfShapeList.get(j);
                    if (sh instanceof XSLFTextShape) {
                        XSLFTextShape shape = (XSLFTextShape) sh;
                        if (shape.getText().length() != 0) {
                            String xml = shape.getXmlObject().xmlText();
                            if (xml.contains("标题 1") || xml.contains("ctrTitle") || xml.contains("type=\"title\"")) {
                                title++;
                            }
                            if (xml.contains("subTitle")) {
                                subTitle++;
                            }
                            if (xml.contains("lvl=\"0\"")) {
                                first_text++;
                            }
                            if (xml.contains("lvl=\"1\"")) {
                                second_text++;
                            }
                            if (xml.contains("lvl=\"2\"")) {
                                third_text++;
                            }
                        }
                    }
                }
                String title_temp = param.get(PptTextBoxPropertiesEnums.TITLE.name()) == null ? "" : param.get(PptTextBoxPropertiesEnums.TITLE.name()).toString();
                String subTitle_temp = param.get(PptTextBoxPropertiesEnums.SUBTITLE.name()) == null ? "" : param.get(PptTextBoxPropertiesEnums.SUBTITLE.name()).toString();
                String first_text_temp = param.get(PptTextBoxPropertiesEnums.FIRST_TEXT.name()) == null ? "" : param.get(PptTextBoxPropertiesEnums.FIRST_TEXT.name()).toString();
                String second_text_temp = param.get(PptTextBoxPropertiesEnums.SECOND_TEXT.name()) == null ? "" : param.get(PptTextBoxPropertiesEnums.SECOND_TEXT.name()).toString();
                String third_text_temp = param.get(PptTextBoxPropertiesEnums.THIRD_TEXT.name()) == null ? "" : param.get(PptTextBoxPropertiesEnums.THIRD_TEXT.name()).toString();
                if (title_temp != null && title_temp.length() != 0) {
                    if (Integer.parseInt(title_temp) == title) {
                        count++;
                        sb.append("标题设置正确" + System.lineSeparator());
                    }
                }
                if (subTitle_temp != null && subTitle_temp.length() != 0) {
                    if (Integer.parseInt(subTitle_temp) == subTitle) {
                        count++;
                        sb.append("子标题设置正确" + System.lineSeparator());
                    }
                }
                if (first_text_temp != null && first_text_temp.length() != 0) {
                    if (Integer.parseInt(first_text_temp) == first_text) {
                        count++;
                        sb.append("一级文本设置正确" + System.lineSeparator());
                    }
                }
                if (second_text_temp != null && second_text_temp.length() != 0) {
                    if (Integer.parseInt(second_text_temp) == second_text) {
                        count++;
                        sb.append("二级文本设置正确" + System.lineSeparator());
                    }
                }
                if (third_text_temp != null && third_text_temp.length() != 0) {
                    if (Integer.parseInt(third_text_temp) == third_text) {
                        count++;
                        sb.append("三级文本设置正确" + System.lineSeparator());
                    }
                }

                // 此规则的总分数
                float temp = Float.parseFloat(score) / sumRule;
                String finalScore = String.valueOf(temp * count / answer_count);
                infoMap.put("score", finalScore);
                infoMap.put("msg", sb.toString());

                return infoMap;

            } else {
                return null;
            }
        }
    }
}