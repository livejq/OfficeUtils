package com.gzmut.office.bean;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * 声音实体类
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-08
 */
@Data
@Accessors(chain = true)
public class Sound {
    /**
     * 声音ID属性
     */
    private String id;

    /**
     * 声音文件名称
     */
    private String name;

    /**
     * 声音文件路径
     */
    private String url;

    /**
     * 声音放映时图标隐藏属性
     */
    private String showWhenStopped;

    /**
     * 判断当前声音对象是否为空
     *
     * @return boolean
     */
    public boolean isNotBlank() {
        return (id != null && name != null && url != null);
    }
}