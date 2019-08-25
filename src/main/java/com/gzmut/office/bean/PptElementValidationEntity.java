package com.gzmut.office.bean;

import com.gzmut.office.enums.ppt.PptTargetEnums;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/**
 * PPT元素校验实体类
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-24
 */
@Data
@Accessors(chain = true)
@AllArgsConstructor
@NoArgsConstructor
public class PptElementValidationEntity {
    /**
     * 校验条ID
     */
    private String id;

    /**
     * 校验元素分值
     */
    private float score;

    /**
     * 校验元素目标
     */
    private PptTargetEnums targetVerify;

    /**
     * 校验 “包含的" 参数JSON
     */
    private String includeParamJson;

    /**
     * 校验 “排除的” 参数JSON
     */
    private String excludeParamJson;

    /**
     * 幻灯片索引
     */
    private int slideIndex;

}