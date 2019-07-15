package poi;

import java.io.Serializable;
import java.util.Map;

public class CheckBean{
    /** 规则编号*/
    public String knowledge;
    /** 定位对象*/
    public Location location;
    /** 参数集合*/
    public int score;

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
     * 定位成员变量 lp,ls
     * 在excel中，lp定位表，ls在lp的基础上定位单元格
     * 在word中，lp定位段落，定位字符串；
     */
    static class Location {
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
