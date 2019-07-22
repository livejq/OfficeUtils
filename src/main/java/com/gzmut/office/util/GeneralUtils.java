package com.gzmut.office.util;

public class GeneralUtils {

    /**
     * 判断数值近似度，误差值在3以内
     * @param value 获取学生提交文件的值
     * @param param 答案参考参数值
     * return 是否近似？true:false
     */
    public static boolean approach(int value, int param){
        //差值
        int diff = 3;
        return value > param - diff && value < param + diff;
    }

}
