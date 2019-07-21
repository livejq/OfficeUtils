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
        String start = "[";
        String end = "]";
        String s1 = "{\"score\":20,\"knowledge\":\"CHECK_FILE_IS_EXIST\",\"param\":{\"fileName\":\"Word素材.docx\"}}";//        JSONArray jsonArray = JSON.parseArray(s);
        String s2 = "{\"score\":20,\"knowledge\":\"CHECK_PAGE\",\"param\":{\"PAGE_SIZE\":\"11907:16839\"}}";
        String s = start + s1+ "," + s2+end;
//        List<CheckBean> checkBeanList = JSONObject.parseArray(jsonArray.toJSONString(), CheckBean.class);
//        checkBeanList.forEach(checkBean -> System.out.println(checkBean.getParam().get("fileName")));
//        WordCorrectImpl wordCorrect = new WordCorrectImpl("E:\\Desktop\\office\\office二级新题(34套)\\26套");
//        Map<String, Object> map = wordCorrect.correctParam(null, checkBeanList.get(0));
//        System.out.println(map);
        // 设置考生文件夹 并初始化判题器
        WordCorrectImpl wordCorrect = new WordCorrectImpl("E:\\\\Desktop\\\\office\\\\office二级新题(34套)\\\\26套");
        List<Map<String, Object>> maps = wordCorrect.correctItem(new File("E:\\\\Desktop\\\\office\\\\office二级新题(34套)\\\\26套\\Word素材.docx"), s);
        System.out.println(maps);
    }
}