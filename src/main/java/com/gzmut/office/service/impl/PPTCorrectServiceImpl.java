package com.gzmut.office.service.impl;

import com.alibaba.fastjson.JSON;
import com.gzmut.office.bean.CheckBean;
import com.gzmut.office.bean.CorrectInfo;
import com.gzmut.office.bean.Location;
import com.gzmut.office.enums.ppt.PPTCorrectEnums;
import com.gzmut.office.service.ICorrect;
import com.gzmut.office.util.PPTCheckUtils;
import com.gzmut.office.util.PPTUtils;
import lombok.Getter;
import lombok.NonNull;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.text.DecimalFormat;
import java.util.*;

/**
 * 主要实现MS Office 2010
 * pptx演示文稿判题的业务
 *
 * @author livejq
 * @since 2019/7/26
 */
@Slf4j
public class PPTCorrectServiceImpl implements ICorrect {

    @Getter
    @Setter
    private String examDir;

    private PPTUtils pptUtils = new PPTUtils();

    private PPTCheckUtils pptCheckUtils = new PPTCheckUtils();

    @Override
    public List<String> getCorrectJson(String num) {
        return null;
    }

    @Override
    public List<CorrectInfo> correct(List<String> correctJson) {

        log.info(System.currentTimeMillis() + ":准备扫描考生文件...");
        List<CorrectInfo> correctInfoList = new LinkedList<>();
        int id = 1;
        int index = 1;
        for (String jsonStr : correctJson) {
            List<Map<String, String>> correctInfo = correctItem(jsonStr);
            float stepScore = 0;
            StringBuilder sb = new StringBuilder();
            for (Map<String, String> info : correctInfo) {
                if (info != null) {
                    for (String key : info.keySet()) {
                        if (key.equals("score")) {
                            System.out.println("规则 " + index++ + " 得分:" + Float.valueOf(info.get(key)));
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

    private CorrectInfo statistics(int id, Float score, String message) {

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
        Map<String, String> param = checkBean.getParam();
        Location location = checkBean.getLocation();
        String score = checkBean.getScore();
        PPTCorrectEnums correctEnums = PPTCorrectEnums.valueOf(checkBean.getKnowledge());

        switch (correctEnums) {
            case CHECK_FILE_IS_EXISTS:
                return null;
            case CHECK_SLIDE_NUMBER:
                return pptCheckUtils.checkSlideNumber(location, param, score, sumRule);
            case CHECK_TEXT_BOX_CONTENT:
                return pptCheckUtils.checkTextBoxContent(location, param, score, sumRule);
            case CHECK_TEXT_BOX_CONTENT_FORMAT:
                return pptCheckUtils.checkTextBoxContentFormat(location, param, score, sumRule);
            case CHECK_BACKGROUND_COLOR:
                return pptCheckUtils.checkBackgroundColor(location, param, score, sumRule);
            case CHECK_PARAGRAPH_STYLE:
                return null;
            case CHECK_LAYOUT_STYLE:
                return pptCheckUtils.checkLayoutStyle(location, param, score, sumRule);
            default:
        }
        return null;
    }
}
