package com.gzmut.office.bean;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * 超链接实体类
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-07
 */
@Data
@Accessors(chain = true)
public class HyperLink {
    /**
     * 超链接ID
     */
    private String id;
    /**
     * 超链接对应行为
     */
    private String linkAction;
    /**
     * 超链接对应文本
     */
    private String linkText;


    /***
     * 判断当前超链接对象是否为空
     *
     * @return boolean
     */
    public boolean isNotBlank() {
        return linkAction != null && linkText != null;
    }
}