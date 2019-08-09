package com.gzmut.office.util;

import com.gzmut.office.enums.ppt.PPTSlidePropertiesEnums;
import org.junit.Test;

/**
 * @author livejq
 * @since 2019/7/22
 */
public class EnumTest {
    @Test
    public void demo01() {
        PPTSlidePropertiesEnums[] pspe = PPTSlidePropertiesEnums.values();
        for(PPTSlidePropertiesEnums s : pspe) {
            System.out.println("输出 name：" + s.name());
        }
    }
}
