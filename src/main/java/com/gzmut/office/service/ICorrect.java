package com.gzmut.office.service;

import com.gzmut.office.bean.CheckBean;
import com.gzmut.office.bean.CorrectInfo;

import java.util.List;
import java.util.Map;

/**
 * Word/Excel/PPT单道大题批改接口
 *
 * @author srl
 * @since 2019-07-10
 */

public interface ICorrect {
    /**
     * 根据小题，判分
     * @param rule 评分规则
     * @return List<Map<String,String>> 判题结果
     */
    List<Map<String,String>> correctItem(String rule);

    /**
     * 根据json数组判分
     * @return 返回一道大题的分数
     */
    List<CorrectInfo> correct(List<String> correctJson);

    /**
     * 根据大题号返回规则
     * @param num 大题号
     * @return 返回规则数组
     */
    List<String> getCorrectJson(String num);

}
