package com.gzmut.office.bean;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.util.Map;

/**
 * @author MXDC
 * @since 2019/7/10
 **/
@Getter
@Setter
@ToString()
public class CheckBean{
    /** 规则编号 */
    private String knowledge;
    /** 定位对象 */
    private Location location;
    /** 参数集合 */
    private Map<String ,Object> param;
    /** 得分 */
    private String score;
    /** 文件名 */
    private String baseFile;
}
