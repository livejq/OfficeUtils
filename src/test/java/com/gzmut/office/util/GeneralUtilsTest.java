package com.gzmut.office.util;

import org.junit.Test;

import static org.junit.Assert.*;

public class GeneralUtilsTest {

    @Test
    public void demo01() {
        boolean result = GeneralUtils.approach(5, 5);
        assertTrue(result);
    }

}