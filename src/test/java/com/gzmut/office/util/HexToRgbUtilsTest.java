package com.gzmut.office.util;

import org.junit.Test;

public class HexToRgbUtilsTest {
    @Test
    public void demo01() {
        String s = HexToRgbUtils.toRgb("FFC0CB");
        System.out.println(s);
    }
}
