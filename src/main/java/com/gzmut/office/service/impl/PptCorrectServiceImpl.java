package com.gzmut.office.service.impl;

import com.alibaba.fastjson.JSON;
import com.gzmut.office.bean.CheckBean;
import com.gzmut.office.bean.CorrectInfo;
import com.gzmut.office.bean.Location;
import com.gzmut.office.enums.ppt.PptCorrectEnums;
import com.gzmut.office.service.ICorrect;
import com.gzmut.office.util.PptCheckUtils;
import com.gzmut.office.util.PptUtils;

import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import lombok.Getter;
import lombok.NonNull;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

/**
 * 主要实现MS Office 2010
 * pptx演示文稿判题的业务
 *
 * @author livejq
 * @since 2019/7/26
 */
@Slf4j
public class PptCorrectServiceImpl implements ICorrect {

    @Getter
    @Setter
    private String examDir;

    private PptUtils pptUtils = new PptUtils();

    private PptCheckUtils pptCheckUtils = new PptCheckUtils();

    @Override
    public List<String> getCorrectJson(String num) {
        return null;
    }

    @Override
    public List<CorrectInfo>  correct(List<String> correctJson) {

        log.info(System.currentTimeMillis() + ":准备扫描考生文件...");
        List<CorrectInfo> correctInfoList = new LinkedList<>();
        int id = 1;
        float stepScore = 0;
        for (String jsonStr : correctJson) {
            List<Map<String, String>> correctInfo = correctItem(jsonStr);
            StringBuilder sb = new StringBuilder();
            for (Map<String, String> info : correctInfo) {
                if (info != null) {
                    for (String key : info.keySet()) {
                        if (key.equals("score")) {
                            stepScore += Float.valueOf(info.get(key));
                        } else if (key.equals("msg")) {
                            sb.append(info.get(key));
                        }
                    }
                }
            }
            correctInfoList.add(statistics(id++, stepScore, sb.toString()));
        }
        log.info(System.currentTimeMillis() + ":扫描完毕！" + System.lineSeparator());
        log.info(System.currentTimeMillis() + "PPT答题总得分：" + stepScore + "分" + System.lineSeparator());

        return correctInfoList;
    }

    @Override
    public List<Map<String, String>> correctItem(@NonNull String rule) {

        List<Map<String, String>> correctInfo = new LinkedList<>();
        List<CheckBean> checkBeanList = JSON.parseArray(rule, CheckBean.class);
        int sumRule = checkBeanList.size();
        checkBeanList.forEach(checkBean -> correctInfo.add(correctKnowledge(checkBean, sumRule)));

        return correctInfo;
    }

    public CorrectInfo statistics(int id, Float score, String message) {

        CorrectInfo info = new CorrectInfo();
        info.setNumber(id);
        info.setScore(score);
        info.setMsg(message);

        return info;
    }

    private Map<String, String> correctKnowledge(CheckBean checkBean, int sumRule) {

        String baseFile = checkBean.getBaseFile();
        String fileName = examDir + File.separator + baseFile;

        // 初始化工具
        if (pptCheckUtils.getPptUtils() == null) {
            pptUtils.initXMLSlideShow(fileName);
            pptCheckUtils.setPptUtils(pptUtils);
        }
        // 由于不是每次过来的文件都相同（即应对多个考场考不同的试卷）
        if (!StringUtils.equals(fileName, pptCheckUtils.getPptUtils().getFileName())) {
            if (!pptUtils.initXMLSlideShow(fileName)) {
                Map<String, String> infoMap = new HashMap<>(2);
                infoMap.put("score", "0.0");
                infoMap.put("msg", "找不到考生文件");

                return infoMap;
            } else {
                // 重新设置相应的判卷工具
                pptCheckUtils.setPptUtils(pptUtils);
            }
        }
        Map<String, Object> param = checkBean.getParam();
        Location location = checkBean.getLocation();
        String score = checkBean.getScore();
        PptCorrectEnums correctEnums = PptCorrectEnums.valueOf(checkBean.getKnowledge());

        switch (correctEnums) {
            case CHECK_SECTION:
                return pptCheckUtils.checkSection(location, param, score, sumRule);
            case CHECK_SWITCH_STYLE:
                return null;
            case CHECK_SLIDE_THEME:
                return pptCheckUtils.checkSlideTheme(location, param, score, sumRule);
            case CHECK_TEXT_HYPERLINK:
                return pptCheckUtils.checkTextHyperlink(location, param, score, sumRule);
            case CHECK_LAYOUT_NAME:
                return pptCheckUtils.checkLayoutName(location, param, score, sumRule);
            case CHECK_TEXT_LEVEL:
                return pptCheckUtils.checkTextLevel(location, param, score, sumRule);
            case CHECK_SLIDE_NUMBER:
                return pptCheckUtils.checkSlideNumber(location, param, score, sumRule);
            case CHECK_TEXT_BOX_CONTENT:
                return pptCheckUtils.checkTextBoxContent(location, param, score, sumRule);
            case CHECK_TEXT_BOX_CONTENT_FORMAT:
                return pptCheckUtils.checkTextBoxContentFormat(location, param, score, sumRule);
            case CHECK_TEXT_ALIGN_STYLE:
                return pptCheckUtils.checkTextAlignStyle(location, param, score, sumRule);
            case CHECK_SLIDE_BACKGROUND_COLOR:
                return pptCheckUtils.checkSlideBackgroundColor(location, param, score, sumRule);
            case CHECK_LAYOUT_STYLE:
                return pptCheckUtils.checkLayoutStyle(location, param, score, sumRule);
            case CHECK_PICTURE_COUNT:
                return null;
            default:
        }

        return null;
    }
}
