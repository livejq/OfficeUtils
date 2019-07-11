package com.gzmut.office.bean;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author MXDC
 * @date 2019/7/10
 **/
public class CheckBean {
    /** 规则编号*/
    public int num;
    /** 定位对象*/
    public Location location;
    /** 参数集合*/
    public Map<String, Object> getParam() {
        return param;
    }
    public void setParam(Map<String, Object> param) {
        this.param = param;
    }
    public Map<String ,Object> param;
    public Location getLocation() {
        return location;
    }
    public void setLocation(Location location) {
        this.location = location;
    }
    public int getNum() {
        return num;
    }
    public void setNum(int num) {
        this.num = num;
    }
    /**
     * 静态内部类
     * 定位类 Location
     * 定位成员变量 lp,ls
     * 在excel中，lp定位表，ls在lp的基础上定位单元格
     * 在word中，lp定位段落，定位字符串；
     */
     static  class Location {
        public String lp,ls;
        public String getLp() {
            return lp;
        }
        public void setLp(String lp) {
            this.lp = lp;
        }
        public String getLs() {
            return ls;
        }
        public void setLs(String ls) {
            this.ls = ls;
        }
    }
}
