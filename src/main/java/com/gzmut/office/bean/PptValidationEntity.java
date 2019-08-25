package com.gzmut.office.bean;

import java.util.List;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * PPT检查信息实体类
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-24
 */
@Data
@Accessors(chain = true)
public class PptValidationEntity {
    /**
     * 唯一ID
     */
    private String id;

    /**
     * PPT 文件名称
     */
    private String fileName;

    /**
     * PPT 元素检验实体类集合
     */
    private List<PptElementValidationEntity> pptElementValidationEntities;
}