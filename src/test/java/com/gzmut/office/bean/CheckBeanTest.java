package com.gzmut.office.bean;

import com.alibaba.fastjson.JSONObject;
import com.google.gson.JsonObject;
import com.uwyn.jhighlight.fastutil.Hash;
import org.junit.Test;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.*;

public class CheckBeanTest {
    @Test
    public void test(){

        CheckBean checkBean = new CheckBean();
        CheckBean.Location location = new CheckBean.Location();
        location.setLp("sheet1");
        location.setLs("A1");
        //checkBean.setLocation(location);
        Map<String,Object> param = new HashMap();
        param.put("loction",location);
        param.put("q",2);
        param.put("b",2);
        param.put("c",2);
        param.put("g",2);
        checkBean.setParam(param);




    }
}