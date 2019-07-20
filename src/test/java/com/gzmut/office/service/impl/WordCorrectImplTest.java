package com.gzmut.office.service.impl;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.gzmut.office.bean.CheckBean;
import com.gzmut.office.util.WordUtils;
import org.junit.Test;

import java.io.File;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.*;

public class WordCorrectImplTest {

    @Test
    public void getJson() {
        String s = "[{\"score\":20,\"knowledge\":\"CHECK_FILE_IS_EXIST\",\"param\":{\"fileName\":\"Word.docx\"}}]";//        JSONArray jsonArray = JSON.parseArray(s);
//        List<CheckBean> checkBeanList = JSONObject.parseArray(jsonArray.toJSONString(), CheckBean.class);
//        checkBeanList.forEach(checkBean -> System.out.println(checkBean.getParam().get("fileName")));
//        WordCorrectImpl wordCorrect = new WordCorrectImpl("E:\\Desktop\\office\\office二级新题(34套)\\26套");
//        Map<String, Object> map = wordCorrect.correctParam(null, checkBeanList.get(0));
//        System.out.println(map);
        WordCorrectImpl wordCorrect = new WordCorrectImpl("E:\\\\Desktop\\\\office\\\\office二级新题(34套)\\\\26套");
        List<Map<String, Object>> maps = wordCorrect.correctItem(null, s);
        System.out.println(maps);
    }
}