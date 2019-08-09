package com.gzmut.office.bean;

import java.util.Objects;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * 字体实体类
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-07
 */
@Data
@Accessors(chain = true)
public class Font {
    /**
     * 中文字体
     */
    private String eaStyle;

    /**
     * 英文字体
     */
    private String latinStyle;

    /**
     * 字体大小
     */
    private String size;

    /**
     * 字体颜色
     */
    private String color;

    /**
     * 字体加粗
     */
    private String bold;

    /**
     * 字体倾斜
     */
    private String tilt;

    /**
     * 下划线
     */
    private String underline;

    /**
     * 删除线
     */
    private String strike;

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        Font font = (Font) o;
        return Objects.equals(eaStyle, font.eaStyle) &&
                Objects.equals(latinStyle, font.latinStyle) &&
                Objects.equals(size, font.size) &&
                Objects.equals(color, font.color) &&
                Objects.equals(bold, font.bold) &&
                Objects.equals(tilt, font.tilt) &&
                Objects.equals(underline, font.underline) &&
                Objects.equals(strike, font.strike);
    }

    @Override
    public int hashCode() {
        return Objects.hash(eaStyle, latinStyle, size, color, bold, tilt, underline, strike);
    }

    /***
     * 判断当前字体对象是否为空
     *
     * @return boolean
     */
    public boolean isNotBlank() {
        return (eaStyle != null || latinStyle != null
                || size != null || color != null
                || bold != null || tilt != null
                || underline != null || strike != null);
    }
}