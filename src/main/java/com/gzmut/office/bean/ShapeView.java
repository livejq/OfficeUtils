package com.gzmut.office.bean;

import com.alibaba.fastjson.annotation.JSONField;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * 形状实体类
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-08
 */
@Data
@Accessors(chain = true)
public class ShapeView {
    /**
     * 形状元素自带ID
     */
    @JSONField(serialize=false)
    private String id;

    /**
     * 形状元素自带name
     */
    private String name;

    /**
     * 形状元素内文本内容
     */
    private String text;

    /**
     * 形状元素轮廓主题
     */
    private String lnVal;

    /**
     * 形状元素填充主题
     */
    private String fillVal;

    /**
     * 形状效果主题
     */
    private String effectVal;

    /**
     * 形状字体主题
     */
    private String fontVal;

    public void setName(String name) {
        name = (name.contains("标题")) ? name : name.split("\\s+\\S*$")[0];
        this.name = name;
    }

    /**
     * 判断当前形状对象是否为空
     *
     * @return boolean
     */
    public boolean isNotBlank() {
        return (id != null && name != null);
    }
}