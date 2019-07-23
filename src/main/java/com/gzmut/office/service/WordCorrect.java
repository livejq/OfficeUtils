package com.gzmut.office.service;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * word判题器接口
 * @author MXDC
 * @date 2019/7/16
 */
public interface WordCorrect {
    /**
     * 根据小题，判分
     * @param file 操作文档
     * @param rule 评分规则
     * @return List<Map<String,object>> 判题结果
     */
    public List<Map<String,Object>> correctItem(File file,String rule);

    /**
     * 根据大题，判分
     * @param textNum 大题号，前端提供
     * @param file 要批改的文件
     * @return 返回一道大题的分数
     */
    public abstract int correct(String textNum, File file);

    /**
     * 根据大题号返回规则
     * @param num 大题号
     * @return 返回规则数组
     */
    public abstract List<String> getJson(String num);
}
