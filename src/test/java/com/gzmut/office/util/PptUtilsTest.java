package com.gzmut.office.util;

import org.junit.Test;

/**
 * @author livejq
 * @date 2019/7/13
 **/
public class PPtUtilsTest {
    @Test
    public void demo01() {
        System.out.println(PPtUtils.getXMLSlideShow("temp/demo01.pptx").getSlides().size());
    }
}
