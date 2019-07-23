package com.gzmut.office.enums;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * office中的常用文件类型
 * @author livejq
 * @date 2019/7/22
 */
@Getter
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public enum OfficeEnum {

    /** Excel xls后缀 */
    XLS(1,".xls"),

    /** Excel xlsx后缀 */
    XLSX(2,".xlsx"),

    /** Word doc后缀 */
    DOC(3,".doc"),

    /** Word docx后缀 */
    DOCX(4,".docx"),

    /** PPt ppt后缀 */
    PPT(5,".ppt"),

    /** PPt pptx后缀 */
    PPTX(6,".pptx");

    private final Integer id;
    private final String suffix;
}
