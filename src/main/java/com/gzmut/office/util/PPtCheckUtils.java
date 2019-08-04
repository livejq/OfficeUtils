package com.gzmut.office.util;

import com.gzmut.office.enums.ppt.PPtSlidePropertiesEnums;
import com.gzmut.office.enums.ppt.PPtTargetEnums;
import com.gzmut.office.enums.ppt.PPtTextBoxPropertiesEnums;
import com.gzmut.office.bean.Location;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.awt.*;
import java.util.HashMap;
import java.util.Map;

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
public class PPtCheckUtils {

    private PPtUtils pptUtils;

    public PPtCheckUtils() {
    }

    public PPtCheckUtils(PPtUtils pptUtils) {
        this.pptUtils = pptUtils;
    }

    public boolean initPptUtils(String fileName) {
        return this.pptUtils.initXMLSlideShow(fileName);
    }

    /**
     * 判断（幻灯片/文本框等）背景颜色是否正确
     *
     * @param location 定位参数
     * @param param    RGB颜色
     * @param score    知识点分数
     * @return Map
     */
    public Map<String, String> checkBackgroundColor(Location location, Map<String, String> param, String score) {

        Map<String, String> infoMap = new HashMap<>(2);
        String[] slides = location.getLp().split(",");

        if (slides != null) {
            for (String slide : slides) {
                PPtTargetEnums target = PPtTargetEnums.valueOf(location.getLo());
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
     * 检查目标元素的数量
     *
     * @param location 定位参数
     * @param param    元素数量
     * @param score    知识点分数
     * @return Map
     */
    public Map<String, String> checkTargetNumber(Location location, Map<String, String> param, String score) {

        Map<String, String> infoMap = new HashMap<>(2);
        PPtTargetEnums target = PPtTargetEnums.valueOf(location.getLo());
        switch (target) {
            case SLIDE:
                if (pptUtils.getSlidesSum() == Integer.parseInt(param.get("COUNT"))) {
                    infoMap.put("score", score);
                    infoMap.put("msg", "幻灯片数量正确!");
                } else {
                    infoMap.put("score", "0.0");
                    infoMap.put("msg", "幻灯片数量错误!");
                }
                return infoMap;
            case TEXT_BOX:
                return null;

            default:

        }

        return null;
    }

    /**
     * 检查幻灯片版式
     *
     * @param location 定位参数
     * @param param    元素数量
     * @param score    知识点分数
     * @return Map
     */
    public Map<String, String> checkLayoutStyle(Location location, Map<String, String> param, String score) {
        Map<String, String> infoMap = new HashMap<>(2);
        String[] slides = location.getLp().split(",");
        int num = 0;
        StringBuilder sb = new StringBuilder("幻灯片布局检查结果， ");
        if (slides != null) {
            for (String slide : slides) {
                PPtTargetEnums target = PPtTargetEnums.valueOf(location.getLo());
                switch (target) {
                    case SLIDE:
                        String[] answer_layout = param.get(PPtSlidePropertiesEnums.LAYOUT_STYLE.getProperty()).split(",");
                        String layout = pptUtils.getSlideLayoutType(pptUtils.getSlideShow().getSlides()
                                .get(Integer.parseInt(slide) - 1)).name();
                        for (int i = 0, len = answer_layout.length; i < len; i++) {
                            if (answer_layout[i].equals(layout)) {
                                num++;
                                sb.append(";演示文稿中第 " + slide + " 张版式正确," + answer_layout[i]);
                            }
                        }
                        infoMap.put("score", score);
                        infoMap.put("msg", sb.toString());

                        return infoMap;
                    default:
                        return null;
                }
            }
        }

        return null;
    }

    /**
     * 检查文本框中的文本格式
     *
     * @param location 定位参数
     * @param param    元素数量
     * @param score    知识点分数
     * @return Map
     */
    public Map<String, String> checkTextBoxContentFormat(Location location, Map<String, String> param, String score) {

        Map<String, String> infoMap = new HashMap<>(2);
        String[] slides = location.getLp().split(",");
        String[] locals = location.getLs().split(",");
        int num = 0;
        StringBuilder sb = new StringBuilder("考卷大纲检查结果， ");
        boolean need_diff_style = false;
        boolean need_diff_size = false;
        boolean need_diff_color = false;
        if (slides != null) {
            for (String slide : slides) {
                if (locals != null) {
                    for (String local : locals) {
                        PPtTargetEnums target = PPtTargetEnums.valueOf(location.getLo());
                        switch (target) {
                            case TEXT_BOX:
                                String answer_content = param.get(PPtTextBoxPropertiesEnums
                                        .TEXT_CONTENT.getProperty());
                                String content = pptUtils.getTextBoxContent(pptUtils.getSlideShow().getSlides()
                                                .get(Integer.parseInt(slide) - 1)
                                        , Integer.parseInt(local) - 1);
                                if (answer_content.equals(content)) {
                                    num++;
                                    sb.append(";文本框中内容正确," + answer_content);
                                }
                                String answer_bold = param.get(PPtTextBoxPropertiesEnums
                                        .BOLD.getProperty());
                                String bold = pptUtils.isTextBoxRawBold(pptUtils.getSlideShow().getSlides()
                                                .get(Integer.parseInt(slide) - 1)
                                        , Integer.parseInt(local) - 1
                                        , Integer.parseInt(location.getLg()) - 1
                                        , Integer.parseInt(location.getLr()) - 1) == true ? "true" : "false";
                                if (answer_bold.equals(bold)) {
                                    num++;
                                    sb.append(";" + "文本框中字体是否粗体," + answer_bold);
                                }
                                String answer_italic = param.get(PPtTextBoxPropertiesEnums
                                        .ITALIC.getProperty());
                                String italic = pptUtils.isTextBoxRawItalic(pptUtils.getSlideShow().getSlides()
                                                .get(Integer.parseInt(slide) - 1)
                                        , Integer.parseInt(local) - 1
                                        , Integer.parseInt(location.getLg()) - 1
                                        , Integer.parseInt(location.getLr()) - 1) == true ? "true" : "false";
                                if (answer_italic.equals(italic)) {
                                    num++;
                                    sb.append(";" + "文本框中字体是否斜体," + answer_bold);
                                }
                                //***********************************************************
                                //***********************************************************
                                need_diff_style = param.get(PPtTextBoxPropertiesEnums
                                        .TEXT_STYLE.getProperty()).substring(0, 1).equals("-");
                                need_diff_size = param.get(PPtTextBoxPropertiesEnums
                                        .TEXT_SIZE.getProperty()).substring(0, 1).equals("-");
                                need_diff_color = param.get(PPtTextBoxPropertiesEnums
                                        .TEXT_COLOR.getProperty()).substring(0, 1).equals("-");

                                // 判断字体样式
                                int len_style = param.get(PPtTextBoxPropertiesEnums
                                        .TEXT_STYLE.getProperty()).length();
                                String answer_style = param.get(PPtTextBoxPropertiesEnums
                                        .TEXT_STYLE.getProperty()).substring(1, len_style);
                                String style = pptUtils.getTextBoxRawFontFamily(pptUtils.getSlideShow().getSlides()
                                                .get(Integer.parseInt(slide) - 1)
                                        , Integer.parseInt(local) - 1
                                        , Integer.parseInt(location.getLg()) - 1
                                        , Integer.parseInt(location.getLr()) - 1);
                                if (need_diff_style) {
                                    if (!answer_style.equals(style)) {
                                        num++;
                                        sb.append(";" + "文本框中字体样式符合要求," + "非" + answer_style);
                                    }
                                } else {
                                    if (answer_style.equals(style)) {
                                        num++;
                                        sb.append(";" + "文本框中字体样式符合要求," + answer_style);
                                    }
                                }

                                // 判断字体大小
                                int len_size = param.get(PPtTextBoxPropertiesEnums
                                        .TEXT_SIZE.getProperty()).length();
                                String answer_size = param.get(PPtTextBoxPropertiesEnums
                                        .TEXT_SIZE.getProperty()).substring(1, len_size);
                                Float size = pptUtils.getTextBoxRawFontSize(pptUtils.getSlideShow().getSlides()
                                                .get(Integer.parseInt(slide) - 1)
                                        , Integer.parseInt(local) - 1
                                        , Integer.parseInt(location.getLg()) - 1
                                        , Integer.parseInt(location.getLr()) - 1);
                                if (need_diff_size) {
                                    if (size != null && Float.parseFloat(answer_size) != size.floatValue()) {
                                        num++;
                                        sb.append(";" + "文本框中字体大小符合要求," + "非" + answer_size);
                                    }
                                } else {
                                    if (Float.parseFloat(answer_size) == size) {
                                        num++;
                                        sb.append(";" + "文本框中字体大小符合要求," + answer_size);
                                    }
                                }

                                // 判断字体颜色
                                int len_color = param.get(PPtTextBoxPropertiesEnums
                                        .TEXT_COLOR.getProperty()).length();
                                String answer_color = param.get(PPtTextBoxPropertiesEnums
                                        .TEXT_COLOR.getProperty()).substring(1, len_color);
                                String color = null;
                                String xml_result = pptUtils.getTextBoxRawFontColor(pptUtils.getSlideShow().getSlides()
                                                .get(Integer.parseInt(slide) - 1)
                                        , Integer.parseInt(local) - 1
                                        , Integer.parseInt(location.getLg()) - 1
                                        , Integer.parseInt(location.getLr()) - 1);
                                if (xml_result != null) {
                                    color = HexToRgbUtils.toRgb(xml_result);
                                }

                                if (need_diff_color) {
                                    if (!answer_color.equals(color)) {
                                        num++;
                                        sb.append(";" + "文本框中字体颜色符合要求," + "非" + answer_color);
                                    }
                                } else {
                                    if (answer_color.equals(color)) {
                                        num++;
                                        sb.append(";" + "文本框中字体颜色符合要求," + answer_color);
                                    }
                                }
                                Float temp = Float.parseFloat(score);
                                String finalScore = String.valueOf(temp / num);
                                infoMap.put("score", finalScore);
                                infoMap.put("msg", sb.toString());

                                return infoMap;
                            default:
                                return null;
                        }
                    }
                }
            }
        }

        return null;
    }
}
