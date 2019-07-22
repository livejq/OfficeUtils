package com.gzmut.office.bean;

import java.util.Map;

/**
 * @author MXDC
 * @date 2019/7/10
 **/
public class CheckBean{
    /** 规则编号 **/
    public String knowledge;
    /** 定位对象 **/
    public Location location;
    /** 参数集合 **/
    public Map<String ,Object> param;
    /** 得分 **/
    public int score;

    public Map<String, Object> getParam() {
        return param;
    }
    public void setParam(Map<String, Object> param) {
        this.param = param;
    }
    public Location getLocation() {
        return location;
    }
    public void setLocation(Location location) {
        this.location = location;
    }
    public String getKnowledge() {
        return knowledge;
    }
    public void setNum(String num) {
        this.knowledge = knowledge;
    }
    public int getScore() {
        return score;
    }
    public void setScore(int score) {
        this.score = score;
    }

    /**
     * 静态内部类
     * 定位类 Location
     * 定位成员变量 lp,ls,lg,lr
     * 在excel中，lp定位表，ls在lp的基础上定位单元格
     * 在word中，lp定位段落，ls定位字符串；
     * 在ppt中，lp定位幻灯片页码，ls 定位元素位置（所有对象解析后根据在幻灯片中的位置会被存入一个数组中），
     * lo定位元素（如幻灯片、文本框等），lg定位段落，lr定位某串字符
     */
   pulic static class Location {
         /** 定位属性 **/

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
        public String getLo() {
            return lo;
        }
        public void setLo(String lo) {
            this.lo = lo;
        }
        public String getLg() {
            return lg;
        }
        public void setLg(String lg) {
            this.lg = lg;
        }
        public String getLr() {
            return lr;
        }
        public void setLr(String lr) {
            this.lr = lr;
        }
    }
}
