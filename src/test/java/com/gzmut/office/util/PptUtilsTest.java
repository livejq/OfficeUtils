package com.gzmut.office.util;

import org.junit.Test;

/**
 * @author livejq
 * @date 2019/7/13
 **/
public class PptUtilsTest {
    @Test
    public void demo01() {
        System.out.println(PptUtils.getSlideFillColor(PptUtils.getXMLSlideShow("temp/demo01.pptx").getSlides().get(0)).getRed());
    }
}
