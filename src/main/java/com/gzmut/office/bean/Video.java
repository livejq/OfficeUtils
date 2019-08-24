package com.gzmut.office.bean;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * 视频实体类
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-08
 */
@Data
@Accessors(chain = true)
public class Video {

    /** 视频ID属性 */
    private String id;

    /** 视频文件名称 */
    private String name;

    /** 视频文件路径 */
    private String url;

    /** 视频封面文件ID属性 */
    private String coverId;

    /** 视频封面文件路径 */
    private String coverUrl;

    /***
     * 判断当前视频对象是否为空
     *
     * @return boolean
     */
    public boolean isNotBlank() {
        return (id != null && name != null
                && url != null && coverId != null
                && coverUrl != null);
    }
}