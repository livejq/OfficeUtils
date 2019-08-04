package com.gzmut.office.bean;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

/**
 * 静态内部类
 * 定位类 Location
 * 定位成员变量 lp,ls,lg,lr
 * 在excel中，lp定位表，ls在lp的基础上定位单元格
 * 在word中，lp定位段落，ls定位字符串；
 * 在ppt中，lp定位幻灯片页码，ls 定位元素位置（所有对象解析后根据在幻灯片中的位置会被存入一个数组中），
 * lo定位元素（如幻灯片、文本框等），lg定位段落，lr定位某串字符
 * @author livejq
 * @since 2019/7/23
 */
@ToString()
public class Location {
    /** 定位属性 */
    @Getter
    @Setter
    private String lp,ls,lo,lg,lr;
}
