package com.gzmut.office.enums;

/**
 * PowerPoint PowerPointConstants
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-09
 */
public class PowerPointConstants {
    /**
     * Prefix in smartArt data resource folder path
     */
    public static final String SMART_ART_DATA_RESOURCE_URL_PREFIX = "/ppt/diagrams/data";

    /**
     * Prefix in smartArt color resource folder path
     */
    public static final String SMART_ART_COLOR_RESOURCE_URL_PREFIX = "/ppt/diagrams/color";

    /**
     * Prefix in smartArt format resource folder path
     */
    public static final String SMART_ART_FORMAT_RESOURCE_URL_PREFIX = "/ppt/diagrams/layout";

    /**
     * Prefix in smartArt style resource folder path
     */
    public static final String SMART_ART_STYLE_RESOURCE_URL_PREFIX = "/ppt/diagrams/quickStyle";

    /**
     * SmartArt element type key
     */
    public static final String SMART_ART_TYPE_KEY = "type";

    /**
     * SmartArt element primary value key
     */
    public static final String SMART_ART_PRIMARY_KEY = "pri";

    /**
     * SmartArt font node key in xml file
     */
    public static final String SMART_ART_TEXT_NODE_TAG = "a:rPr";

    /**
     * SmartArt latin node key in xml file
     */
    public static final String SMART_ART_TEXT_LATIN_STYLE_XPATH = "//a:latin/@typeface";

    /**
     * SmartArt EA node key in xml file
     */
    public static final String SMART_ART_TEXT_EA_STYLE_XPATH = "//a:ea/@typeface";

    /**
     * SmartArt font size key in xml file
     */
    public static final String SMART_ART_TEXT_SIZE_NAME = "@sz";

    /**
     * SmartArt font bold key in xml file
     */
    public static final String SMART_ART_TEXT_BOLD_NAME = "@b";

    /**
     * SmartArt font tile key in xml file
     */
    public static final String SMART_ART_TEXT_TILT_NAME = "@i";

    /**
     * SmartArt font underline key in xml file
     */
    public static final String SMART_ART_TEXT_UNDERLINE_NAME = "@u";

    /**
     * SmartArt font strike key in xml file
     */
    public static final String SMART_ART_TEXT_STRIKE_NAME = "@strike";

    /**
     * SmartArt font color value where in xml path
     */
    public static final String SMART_ART_TEXT_COLOR_XPATH = "//a:srgbClr/@val";

    /**
     * SmartArt photo value where in xml path
     */
    public static final String SMART_ART_PHOTO_XPATH = "//a:blip//@r:embed";

    /**
     * SmartArt hyperlink tag where in xml path
     */
    public static final String SMART_ART_HYPERLINK_XML_TAG = "a:hlinkClick";

    /**
     * SmartArt hyperlink action value where in xml path
     */
    public static final String SMART_ART_HYPERLINK_ACTION_NAME = "@action";

    /**
     * SmartArt hyperlink id where in xml path
     */
    public static final String SMART_ART_HYPERLINK_RID_NAME = "@r:id";

    /**
     * PowerPoint file slide xml crucial node xml tag
     */
    public static final String CRUCIAL_NODE_XML_TAG = "p:cNvPr";

    /**
     * ShapeView outline's color value where in xml path
     */
    public static final String SHAPE_OUTLINE_COLOR_VALUE_XPATH = "../..//a:lnRef/a:schemeClr/@val";

    /**
     * ShapeView fill color value where in xml path
     */
    public static final String SHAPE_FILL_COLOR_VALUE_XPATH = "../..//a:fillRef/a:schemeClr/@val";

    /**
     * ShapeView effect color value where in xml path
     */
    public static final String SHAPE_EFFECT_COLOR_VALUE_XPATH = "../..//a:effectRef/a:schemeClr/@val";

    /**
     * ShapeView font color value where in xml path
     */
    public static final String SHAPE_FONT_COLOR_VALUE_XPATH = "../..//a:fontRef/a:schemeClr/@val";

    /**
     * PowerPoint xml file id key
     */
    public static final String ID_ATTRIBUTE_KEY = "@id";

    /**
     * PowerPoint xml file name key
     */
    public static final String NAME_ATTRIBUTE_KEY = "@name";

    /**
     * PowerPoint file slide resource url prefix
     */
    public static final String SLIDE_RESOURCE_URL_PREFIX = "/ppt/slides/slide";

    /**
     * PowerPoint file slide xml video tag
     */
    public static final String SLIDE_VIDEO_XML_TAG = "a:videoFile";

    /**
     * PowerPoint file slide xml video resource id name
     */
    public static final String SLIDE_RESOURCE_LINK_ID_XML_ATTRIBUTE = "@r:link";

    /**
     * PowerPoint file slide xml sound file tag
     */
    public static final String SLIDE_SOUND_FILE_XML_TAG = "a:audioFile";

    /**
     * PowerPoint file slide xml sound media tag
     */
    public static final String SLIDE_SOUND_NODE_XML_TAG = "p:cMediaNode";



    /**
     * PowerPoint file slide xml sound show when stop xml attribute name
     */
    public static final String SLIDE_SOUND_SHOW_WHEN_STOP_XML_ATTRIBUTE = "@showWhenStopped";


}