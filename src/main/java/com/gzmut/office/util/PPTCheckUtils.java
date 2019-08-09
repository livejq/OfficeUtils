package com.gzmut.office.util;

import com.gzmut.office.enums.ppt.PPTSlidePropertiesEnums;
import com.gzmut.office.enums.ppt.PPTTargetEnums;
import com.gzmut.office.enums.ppt.PPTTextBoxPropertiesEnums;
import com.gzmut.office.bean.Location;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;

import java.awt.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;

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
public class PPTCheckUtils {

    private PPTUtils pptUtils;

    public PPTCheckUtils() {
    }

    public PPTCheckUtils(PPTUtils pptUtils) {
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
    public Map<String, String> checkBackgroundColor(Location location, Map<String, String> param, String score, int sumRule) {

        Map<String, String> infoMap = new HashMap<>(2);
        String[] slides = location.getLp().split(",");

        if (slides != null) {
            for (String slide : slides) {
                PPTTargetEnums target = PPTTargetEnums.valueOf(location.getLo());
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

                        String backgroundColor = param.get("BACKGROUND_COLOR");
                        String[] answer_color = backgroundColor.split(",");

                        if (red == Integer.parseInt(answer_color[0]) && green == Integer.parseInt(answer_color[1]) && blue == Integer.parseInt(answer_color[2])) {
                            infoMap.put("score", score);
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
    public Map<String, String> checkSlideNumber(Location location, Map<String, String> param, String score, int sumRule) {

        Map<String, String> infoMap = new HashMap<>(2);
        int answer_sumPage = Integer.parseInt(param.get("COUNT"));
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
     * 检查幻灯片版式
     *
     * @param location 定位参数
     * @param param    元素数量
     * @param score    小题总分
     * @param sumRule  小题规则总数
     * @return Map
     */
    public Map<String, String> checkLayoutStyle(Location location, Map<String, String> param, String score, int sumRule) {
        Map<String, String> infoMap = new HashMap<>(2);
        String slide = location.getLp();
        int num = 0;
        StringBuilder sb = new StringBuilder("[幻灯片布局检查结果]");
        if (slide != null || slide.length() != 0) {
            String answer_layout = param.get(PPTSlidePropertiesEnums.LAYOUT_STYLE.getProperty());
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
     * 检查目标元素的数量
     *
     * @param location 定位参数
     * @param param    元素数量
     * @param score    小题总分
     * @param sumRule  小题规则总数
     * @return Map
     */
    public Map<String, String> checkTextBoxContent(Location location, Map<String, String> param, String score, int sumRule) {

        Map<String, String> infoMap = new HashMap<>(2);
        StringBuilder sb = new StringBuilder("[ppt文本内容检查结果]" + System.lineSeparator());
        int num = 0;
        int textBoxCount = 0;
        int page = Integer.parseInt(location.getLp()) - 1;
        String[] answer_content = param.get(PPTTextBoxPropertiesEnums
                .TEXT_CONTENT.getProperty()).split("&#@@#&");

        List<XSLFShape> xslfShapeList = pptUtils.getShape(page);
        sb.append("第 " + (page + 1) + " 张幻灯片:");

        if (xslfShapeList != null) {
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

    /**
     * 检查文本框中的文本格式
     *
     * @param location 定位参数
     * @param param    元素数量
     * @param score    小题总分
     * @param sumRule  小题规则总数
     * @return Map
     */
    public Map<String, String> checkTextBoxContentFormat(Location location, Map<String, String> param, String score, int sumRule) {

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
                sb.append("第 " + page + " 幻灯片:");
                List<XSLFShape> xslfShapeList = slides.get(page - 1).getShapes();
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
                        Object strColor = param.get(PPTTextBoxPropertiesEnums
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
                        Object strStyle = param.get(PPTTextBoxPropertiesEnums
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
                        Object strSize = param.get(PPTTextBoxPropertiesEnums
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
                    List<XSLFShape> xslfShapeList = slides.get(index - 1).getShapes();
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
                            String strColor = param.get(PPTTextBoxPropertiesEnums.TEXT_COLOR.getProperty());
                            String answer_hexColor = (strColor == null ? "" : strColor.toString());
                            if (hexColor.contains(answer_hexColor)) {
                                num++;
                                if (answer_hexColor.length() == 0) {
                                    answer_hexColor = "默认值";
                                }
                                sb.append("文本框中字体颜色正确:" + answer_hexColor + System.lineSeparator());
                            }
                            // 判断字体样式
                            String fontStyle = null;
                            fontStyle = pptUtils.getTextRunFontStyle(rList, randRaw);

                            Object strStyle = param.get(PPTTextBoxPropertiesEnums
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

                            Object strSize = param.get(PPTTextBoxPropertiesEnums
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
}