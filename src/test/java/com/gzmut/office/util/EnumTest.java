package com.gzmut.office.util;

import com.gzmut.office.enums.ppt.PPtSlidePropertiesEnums;
import org.junit.Test;

/**
 * @author livejq
 * @date 2019/7/22
 **/
public class EnumTest {
    @Test
    public void demo01() {
        PPtSlidePropertiesEnums[] pspe = PPtSlidePropertiesEnums.values();
        for(PPtSlidePropertiesEnums s : pspe) {
            System.out.println("输出 name：" + s.name());
        }
    }
}
