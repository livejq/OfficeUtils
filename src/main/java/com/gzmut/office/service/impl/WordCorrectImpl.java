package com.gzmut.office.service.impl;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.gzmut.office.bean.CheckBean;
import com.gzmut.office.enums.word.WordCorrectEnums;
import com.gzmut.office.service.WordCorrect;
import com.gzmut.office.util.WordUtils;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * word判断题器实现类
 * 使用判题器类时要，初始化考试文件夹 examDir字段
 * @author MXDC
 * @date 2019/7/16
 **/
@Slf4j
public class WordCorrectImpl implements WordCorrect {


    /**
     * 考试文件夹，全路径
     */
    private String examDir;

    public WordCorrectImpl(String examDir) {
        this.examDir = examDir;
        WordUtils.examDir = examDir;
    }

    public String getExamDir() {
        return examDir;
    }

    public void setExamDir(String examDir) {
        this.examDir = examDir;
    }


    /**
     * 小题判题器
     *  1.解析判题规则
     *  2.获取判题枚举类
     *  3.根据判题枚举类判题，并返回判题结果信息
     * @param file 操作文档
     * @param rule 评分规则
     * @return Map<String, Object></> 判题结果
     */
    @Override
    public List<Map<String, Object>> correctItem(File file, String rule) {
        try {
            WordUtils.setDocment(file.getCanonicalPath());
        } catch (IOException e) {
            e.printStackTrace();
            log.error("文档读取错误");
        }
        // correctInfo 小题判题结果信息
        List<Map<String, Object>> correctInfo = new LinkedList<>();
        // 1.解析判题规则
        JSONArray jsonArray = JSON.parseArray(rule);
        List<CheckBean> checkBeanList = JSONObject.parseArray(jsonArray.toJSONString(), CheckBean.class);
        checkBeanList.forEach(checkBean ->correctInfo.add(correctParam(checkBean)));
        // 2.获取判题枚举类
        // 3.判题并返回判题信息
        return correctInfo;
    }

    /**
     * 根据ckeckBean 按得分点(参数)给小题判分
     * @param checkBean 评分规则
     * @return Map<String, Object></> 判题结果
     */
    public Map<String, Object> correctParam(CheckBean checkBean){
        // 判题规则枚举类名称
        String knowledge;
        // 定位信息
        CheckBean.Location location = checkBean.getLocation();
        // 判题参数
        Map<String, Object> param = checkBean.getParam();
        // 题目分数
        int score = checkBean.getScore();
        knowledge = checkBean.getKnowledge();
        // 判题枚举类
        WordCorrectEnums wordCorrectEnums = WordCorrectEnums.valueOf(knowledge);
        // 判题
        switch (wordCorrectEnums){
            case CHECK_FILE_IS_EXIST:
                return WordUtils.checkFileIsExist(score,param);
            case CHECK_PAGE:
                return WordUtils.checkPage(score,param);

            default:
        }
        return null;
    }

    @Override
    public int correct(String textNum, File file) {
        return 0;
    }

    @Override
    public List<String> getJson(String num) {
        return null;
    }
}
