package com.gzmut.office.service.impl;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.gzmut.office.bean.PptElementValidationEntity;
import com.gzmut.office.bean.PptValidationEntity;
import com.gzmut.office.bean.ShapeView;
import com.gzmut.office.bean.Sound;
import com.gzmut.office.bean.Video;
import com.gzmut.office.enums.CheckInfoResult;
import com.gzmut.office.enums.CheckStatus;
import com.gzmut.office.enums.ParamMatchEnum;
import com.gzmut.office.enums.PowerPointConstants;
import com.gzmut.office.enums.ppt.PptCorrectEnums;
import com.gzmut.office.enums.ppt.PptTargetEnums;
import com.gzmut.office.util.EnumUtils;
import com.gzmut.office.util.PptUtils;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.junit.Test;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class PptCorrectServiceImplTest {
    /**
     * JUnit test class resource folder path
     */
    private String resourcePath = getClass().getResource("/").getPath();

    /**
     * JUnit test target include param JSON file path
     */
    private String smartArtIncludeParamJson = resourcePath + "ppt_smartArt_includeParamJson.json";

    private String videoIncludeParamJson = resourcePath + "ppt_video_includeParamJson.json";

    private String soundIncludeParamJson = resourcePath + "ppt_sound_includeParamJson.json";

    private String shapeIncludeParamJson = resourcePath + "ppt_shape_includeParamJson.json";


    /**
     * JUnit test target exclude param JSON file path
     */
    private String smartArtExcludeParamJson = resourcePath + "ppt_smartArt_excludeParamJson.json";

    private String videoExcludeParamJson = resourcePath + "ppt_video_excludeParamJson.json";

    private String soundExcludeParamJson = resourcePath + "ppt_sound_excludeParamJson.json";

    private String shapeExcludeParamJson = resourcePath + "ppt_shape_excludeParamJson.json";

    /**
     * JUnit test target PowerPoint file path
     */
    private String fileName = resourcePath + "Test.pptx";

    /**
     * Time format
     */
    private SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss:SSS");

    private PptUtils pptUtils = new PptUtils();

    @Test
    public void ruleDemo() throws IOException {
        long startTime = System.currentTimeMillis();
        System.out.println("初始化判题规则中——启动时间：" + dateFormat.format(startTime));

        List<CheckInfoResult> checkInfoResults;
        List<PptElementValidationEntity> pptElementValidationEntities = new ArrayList<>();

        // SmartArt 元素规则类
        PptElementValidationEntity pptSmartArtValidationEntity = new PptElementValidationEntity("01", 3.0f, PptTargetEnums.SMART_ART, readJsonData(smartArtIncludeParamJson), readJsonData(smartArtExcludeParamJson), 3);
        // Video 元素规则类
        PptElementValidationEntity pptVideoValidationEntity = new PptElementValidationEntity("02", 3.0f, PptTargetEnums.VIDEO, readJsonData(videoIncludeParamJson), readJsonData(videoExcludeParamJson), 7);
        // Sound 元素规则类
        PptElementValidationEntity pptSoundValidationEntity = new PptElementValidationEntity("03", 3.0f, PptTargetEnums.SOUND, readJsonData(soundIncludeParamJson), readJsonData(soundExcludeParamJson), 5);
        // Shape 元素规则类
        PptElementValidationEntity pptShapeValidationEntity = new PptElementValidationEntity("04", 3.0f, PptTargetEnums.SHAPE, readJsonData(shapeIncludeParamJson), readJsonData(shapeExcludeParamJson), 2);

        pptElementValidationEntities.add(pptSmartArtValidationEntity);
        pptElementValidationEntities.add(pptVideoValidationEntity);
        pptElementValidationEntities.add(pptSoundValidationEntity);
        pptElementValidationEntities.add(pptShapeValidationEntity);

        PptValidationEntity pptValidationEntity = new PptValidationEntity();
        pptValidationEntity.setId("01")
                .setFileName("PPT.pptx")
                .setPptElementValidationEntities(pptElementValidationEntities);

        long endTime = System.currentTimeMillis();
        System.out.println("初始化判题规则完成——完成时间：" + dateFormat.format(endTime) + "——耗时：" + (endTime - startTime) + "ms");
        checkInfoResults = checkPpt(fileName, pptValidationEntity);
        checkInfoResults.forEach(
                checkInfoResult -> System.out.println(System.lineSeparator() + "当前校验文件——" + fileName + System.lineSeparator() + "校验结果：" + checkInfoResult.getStatus().getDesc() + "——得分：" + checkInfoResult.getScore() + "——详细信息：" + checkInfoResult.getMessage())
        );
    }

    /**
     * PPT文件校验方法（精度：单张幻灯片元素级）
     *
     * @param elementValidationEntity 所需校验的元素实体类
     * @return com.gzmut.office.enums.CheckInfoResult
     */
    public CheckInfoResult checkPptElement(PptElementValidationEntity elementValidationEntity) {
        if (pptUtils.getSlideShow() == null) return CheckInfoResult.exception(0, "找不到考生文件");

        PptTargetEnums target = elementValidationEntity.getTargetVerify();
        System.out.println("正在校验" + target.getDesc() + "元素");
        switch (target) {
            case SMART_ART:
                return checkSmartArt(pptUtils.getSlideShow(), elementValidationEntity);
            case VIDEO:
                return checkVideo(pptUtils.getSlideShow(), elementValidationEntity);
            case SOUND:
                return checkSound(pptUtils.getSlideShow(), elementValidationEntity);
            case SHAPE:
                return checkShape(pptUtils.getSlideShow(), elementValidationEntity);
            default:
                return CheckInfoResult.exception(0, "解析目标不存在");
        }
    }

    /**
     * 针对PPT文件校验实体类对其中元素实体类进行遍历校验（精度：PPT文件级）
     *
     * @param filePath            所需检测考生文件路径
     * @param pptValidationEntity PPT校验实体类
     * @return java.util.List<com.gzmut.office.enums.CheckInfoResult>
     */
    public List<CheckInfoResult> checkPpt(String filePath, PptValidationEntity pptValidationEntity) {
        long startTime = System.currentTimeMillis();
        System.out.println("开始扫描考生文件——启动时间：" + dateFormat.format(startTime));
        List<CheckInfoResult> checkInfoResults = new ArrayList<>();
        if (!pptUtils.initXMLSlideShow(filePath)) {
            checkInfoResults.add(CheckInfoResult.exception(0, "找不到考生文件"));
        } else {
            List<PptElementValidationEntity> pptElementValidationEntities = pptValidationEntity.getPptElementValidationEntities();
            pptElementValidationEntities.forEach(elementValidationEntity -> checkInfoResults.add(checkPptElement(elementValidationEntity)));
        }
        long endTime = System.currentTimeMillis();
        System.out.println("扫描完成——完成时间：" + dateFormat.format(endTime) + "——耗时：" + (endTime - startTime) + "ms");
        return checkInfoResults;
    }

    /**
     * 针对PPT元素实体校验类对幻灯片对象中具体幻灯片页进行SmartArt元素校验
     *
     * @param xmlSlideShow            所需校验的幻灯片对象
     * @param elementValidationEntity PPT元素实体校验类
     * @return com.gzmut.office.enums.CheckInfoResult
     */
    public CheckInfoResult checkSmartArt(XMLSlideShow xmlSlideShow, PptElementValidationEntity elementValidationEntity) {
        if (elementValidationEntity.getSlideIndex() > xmlSlideShow.getSlides().size()) {
            return CheckInfoResult.wrong(0, "无法读取第" + (elementValidationEntity.getSlideIndex() + 1) + "张幻灯片");
        }

        XSLFSlide slide = xmlSlideShow.getSlides().get(elementValidationEntity.getSlideIndex());

        Map<String, Object> includeParams = JSON.parseObject(elementValidationEntity.getIncludeParamJson());
        Map<String, Object> excludeParams = JSON.parseObject(elementValidationEntity.getExcludeParamJson());

        CheckInfoResult checkInfoResult = new CheckInfoResult(CheckStatus.CORRECT, 0, System.lineSeparator(), elementValidationEntity.getTargetVerify());
        PptCorrectEnums targetEnum;
        float averageScore = elementValidationEntity.getScore() / (float) (includeParams.size() + excludeParams.size());

        System.out.println("开始" + ParamMatchEnum.INCLUDE.getDesc() + "校验");
        try {
            for (String target : includeParams.keySet()) {
                targetEnum = Objects.requireNonNull(EnumUtils.getEnumByName(PptCorrectEnums.class, target));
                switch (targetEnum) {
                    case TEXT_CONTENT:
                        checkInfoResult = compareTextContent(checkInfoResult, includeParams.get(target), pptUtils.getSmartArtSpliceText(slide), averageScore, targetEnum.getDesc(), ParamMatchEnum.INCLUDE);
                        break;
                    case FORMAT:
                        checkInfoResult = compareSingleLayerParam(checkInfoResult, includeParams.get(target), pptUtils.getSmartArtLayoutAttributes(slide, PowerPointConstants.SMART_ART_FORMAT_RESOURCE_URL_PREFIX), averageScore, targetEnum.getDesc(), ParamMatchEnum.INCLUDE);
                        break;
                    case COLOR:
                        checkInfoResult = compareSingleLayerParam(checkInfoResult, includeParams.get(target), pptUtils.getSmartArtLayoutAttributes(slide, PowerPointConstants.SMART_ART_COLOR_RESOURCE_URL_PREFIX), averageScore, targetEnum.getDesc(), ParamMatchEnum.INCLUDE);
                        break;
                    case STYLE:
                        checkInfoResult = compareSingleLayerParam(checkInfoResult, includeParams.get(target), pptUtils.getSmartArtLayoutAttributes(slide, PowerPointConstants.SMART_ART_STYLE_RESOURCE_URL_PREFIX), averageScore, targetEnum.getDesc(), ParamMatchEnum.INCLUDE);
                        break;
                    case FONT:
                        checkInfoResult = compareMultiLayerParam(checkInfoResult, includeParams.get(target), pptUtils.getSmartArtFontStyle(slide), averageScore, targetEnum.getDesc(), ParamMatchEnum.INCLUDE);
                        break;
                    case HYPERLINK:
                        checkInfoResult = compareMultiLayerParam(checkInfoResult, includeParams.get(target), pptUtils.getSmartArtHyperLink(slide), averageScore, targetEnum.getDesc(), ParamMatchEnum.INCLUDE);
                        break;
                    default:
                        break;
                }
            }
            System.out.println("校验完成");

            System.out.println("开始" + ParamMatchEnum.EXCLUDE.getDesc() + "校验");
            for (String target : excludeParams.keySet()) {
                targetEnum = Objects.requireNonNull(EnumUtils.getEnumByName(PptCorrectEnums.class, target));
                switch (targetEnum) {
                    case TEXT_CONTENT:
                        checkInfoResult = compareTextContent(checkInfoResult, excludeParams.get(target), pptUtils.getSmartArtSpliceText(slide), averageScore, targetEnum.getDesc(), ParamMatchEnum.EXCLUDE);
                        break;
                    case FORMAT:
                        checkInfoResult = compareSingleLayerParam(checkInfoResult, excludeParams.get(target), pptUtils.getSmartArtLayoutAttributes(slide, PowerPointConstants.SMART_ART_FORMAT_RESOURCE_URL_PREFIX), averageScore, targetEnum.getDesc(), ParamMatchEnum.EXCLUDE);
                        break;
                    case COLOR:
                        checkInfoResult = compareSingleLayerParam(checkInfoResult, excludeParams.get(target), pptUtils.getSmartArtLayoutAttributes(slide, PowerPointConstants.SMART_ART_COLOR_RESOURCE_URL_PREFIX), averageScore, targetEnum.getDesc(), ParamMatchEnum.EXCLUDE);
                        break;
                    case STYLE:
                        checkInfoResult = compareSingleLayerParam(checkInfoResult, excludeParams.get(target), pptUtils.getSmartArtLayoutAttributes(slide, PowerPointConstants.SMART_ART_STYLE_RESOURCE_URL_PREFIX), averageScore, targetEnum.getDesc(), ParamMatchEnum.EXCLUDE);
                        break;
                    case FONT:
                        checkInfoResult = compareMultiLayerParam(checkInfoResult, excludeParams.get(target), pptUtils.getSmartArtFontStyle(slide), averageScore, targetEnum.getDesc(), ParamMatchEnum.EXCLUDE);
                        break;
                    case HYPERLINK:
                        checkInfoResult = compareMultiLayerParam(checkInfoResult, excludeParams.get(target), pptUtils.getSmartArtHyperLink(slide), averageScore, targetEnum.getDesc(), ParamMatchEnum.EXCLUDE);
                        break;
                    default:
                        break;
                }
            }
            System.out.println("校验完成");
        } catch (Exception e) {
            e.printStackTrace();
            return CheckInfoResult.exception(0, elementValidationEntity + System.lineSeparator() + "解析发生异常:" + e.getMessage());
        }
        return checkInfoResult;
    }

    /**
     * 针对PPT元素实体校验类对幻灯片对象中具体幻灯片页进行Video元素校验
     *
     * @param xmlSlideShow            所需校验的幻灯片对象
     * @param elementValidationEntity PPT元素实体校验类
     * @return com.gzmut.office.enums.CheckInfoResult
     */
    public CheckInfoResult checkVideo(XMLSlideShow xmlSlideShow, PptElementValidationEntity elementValidationEntity) {
        if (elementValidationEntity.getSlideIndex() > xmlSlideShow.getSlides().size()) {
            return CheckInfoResult.wrong(0, "无法读取第" + (elementValidationEntity.getSlideIndex() + 1) + "张幻灯片");
        }

        XSLFSlide slide = xmlSlideShow.getSlides().get(elementValidationEntity.getSlideIndex());

        Map<String, Object> includeParams = JSON.parseObject(elementValidationEntity.getIncludeParamJson());
        Map<String, Object> excludeParams = JSON.parseObject(elementValidationEntity.getExcludeParamJson());

        CheckInfoResult checkInfoResult = new CheckInfoResult(CheckStatus.CORRECT, 0, System.lineSeparator(), elementValidationEntity.getTargetVerify());
        PptCorrectEnums targetEnum;
        float averageScore = elementValidationEntity.getScore() / (float) (includeParams.size() + excludeParams.size());

        System.out.println("开始" + ParamMatchEnum.INCLUDE.getDesc() + "校验");

        try {
            for (String target : includeParams.keySet()) {
                targetEnum = Objects.requireNonNull(EnumUtils.getEnumByName(PptCorrectEnums.class, target));
                switch (targetEnum) {
                    case NAME: {
                        List<String> correctSoundName = JSON.parseArray(String.valueOf(excludeParams.get(target)), Video.class)
                                .stream()
                                .map(Video::getName)
                                .collect(Collectors.toList());
                        List<String> videoNameParam = pptUtils.getVideo(slide)
                                .stream()
                                .map(Video::getName)
                                .collect(Collectors.toList());
                        checkInfoResult = compareSingleLayerParam(checkInfoResult, correctSoundName, videoNameParam, averageScore, targetEnum.getDesc(), ParamMatchEnum.INCLUDE);
                        break;
                    }
                    default:
                        break;
                }
            }
            System.out.println("校验完成");

            System.out.println("开始" + ParamMatchEnum.EXCLUDE.getDesc() + "校验");
            for (String target : excludeParams.keySet()) {
                targetEnum = Objects.requireNonNull(EnumUtils.getEnumByName(PptCorrectEnums.class, target));
                switch (targetEnum) {
                    case NAME: {
                        List<String> correctSoundName = JSON.parseArray(String.valueOf(excludeParams.get(target)), Video.class)
                                .stream()
                                .map(Video::getName)
                                .collect(Collectors.toList());
                        List<String> videoNameParam = pptUtils.getVideo(slide)
                                .stream()
                                .map(Video::getName)
                                .collect(Collectors.toList());
                        checkInfoResult = compareSingleLayerParam(checkInfoResult, correctSoundName, videoNameParam, averageScore, targetEnum.getDesc(), ParamMatchEnum.EXCLUDE);
                        break;
                    }
                    default:
                        break;
                }
            }
            System.out.println("校验完成");
        } catch (Exception e) {
            e.printStackTrace();
            return CheckInfoResult.exception(0, elementValidationEntity + "解析发生异常:" + e.getMessage());
        }
        return checkInfoResult;
    }

    /**
     * 针对PPT元素实体校验类对幻灯片对象中具体幻灯片页进行Sound元素校验
     *
     * @param xmlSlideShow            所需校验的幻灯片对象
     * @param elementValidationEntity PPT元素实体校验类
     * @return com.gzmut.office.enums.CheckInfoResult
     */
    public CheckInfoResult checkSound(XMLSlideShow xmlSlideShow, PptElementValidationEntity elementValidationEntity) {
        if (elementValidationEntity.getSlideIndex() > xmlSlideShow.getSlides().size()) {
            return CheckInfoResult.wrong(0, "无法读取第" + (elementValidationEntity.getSlideIndex() + 1) + "张幻灯片");
        }

        XSLFSlide slide = xmlSlideShow.getSlides().get(elementValidationEntity.getSlideIndex());

        Map<String, Object> includeParams = JSON.parseObject(elementValidationEntity.getIncludeParamJson());
        Map<String, Object> excludeParams = JSON.parseObject(elementValidationEntity.getExcludeParamJson());


        CheckInfoResult checkInfoResult = new CheckInfoResult(CheckStatus.CORRECT, 0, System.lineSeparator(), elementValidationEntity.getTargetVerify());
        PptCorrectEnums targetEnum;
        float averageScore = elementValidationEntity.getScore() / (float) (includeParams.size() + excludeParams.size());

        System.out.println("开始" + ParamMatchEnum.INCLUDE.getDesc() + "校验");

        try {
            for (String target : includeParams.keySet()) {
                targetEnum = Objects.requireNonNull(EnumUtils.getEnumByName(PptCorrectEnums.class, target));
                switch (targetEnum) {
                    case NAME: {
                        List<String> correctSoundName = JSON.parseArray(String.valueOf(includeParams.get(target)), Sound.class)
                                .stream()
                                .map(Sound::getName)
                                .collect(Collectors.toList());
                        List<String> soundNameParam = pptUtils.getSound(slide)
                                .stream()
                                .map(Sound::getName)
                                .collect(Collectors.toList());
                        checkInfoResult = compareSingleLayerParam(checkInfoResult, correctSoundName, soundNameParam, averageScore, targetEnum.getDesc(), ParamMatchEnum.INCLUDE);
                        break;
                    }
                    case ATTRIBUTE: {
                        List<String> correctSoundAttr = JSON.parseArray(String.valueOf(includeParams.get(target)), Sound.class)
                                .stream()
                                .map(Sound::getShowWhenStopped)
                                .collect(Collectors.toList());
                        List<String> soundAttrParam = pptUtils.getSound(slide)
                                .stream()
                                .map(Sound::getShowWhenStopped)
                                .collect(Collectors.toList());
                        checkInfoResult = compareSingleLayerParam(checkInfoResult, correctSoundAttr, soundAttrParam, averageScore, targetEnum.getDesc(), ParamMatchEnum.INCLUDE);
                        break;
                    }
                    default:
                        break;
                }
            }
            System.out.println("校验完成");

            System.out.println("开始" + ParamMatchEnum.EXCLUDE.getDesc() + "校验");
            for (String target : excludeParams.keySet()) {
                targetEnum = Objects.requireNonNull(EnumUtils.getEnumByName(PptCorrectEnums.class, target));
                switch (targetEnum) {
                    case NAME: {
                        List<String> correctSoundName = JSON.parseArray(String.valueOf(excludeParams.get(target)), Sound.class)
                                .stream()
                                .map(Sound::getName)
                                .collect(Collectors.toList());
                        List<String> soundNameParam = pptUtils.getSound(slide)
                                .stream()
                                .map(Sound::getName)
                                .collect(Collectors.toList());
                        checkInfoResult = compareSingleLayerParam(checkInfoResult, correctSoundName, soundNameParam, averageScore, targetEnum.getDesc(), ParamMatchEnum.EXCLUDE);
                        break;
                    }
                    case ATTRIBUTE: {
                        List<String> correctSoundAttr = JSON.parseArray(String.valueOf(excludeParams.get(target)), Sound.class)
                                .stream()
                                .map(Sound::getShowWhenStopped)
                                .collect(Collectors.toList());
                        List<String> soundAttrParam = pptUtils.getSound(slide)
                                .stream()
                                .map(Sound::getShowWhenStopped)
                                .collect(Collectors.toList());
                        checkInfoResult = compareSingleLayerParam(checkInfoResult, correctSoundAttr, soundAttrParam, averageScore, targetEnum.getDesc(), ParamMatchEnum.EXCLUDE);
                        break;
                    }
                    default:
                        break;
                }
            }
            System.out.println("校验完成");
        } catch (Exception e) {
            e.printStackTrace();
            return CheckInfoResult.exception(0, elementValidationEntity + "解析发生异常:" + e.getMessage());
        }
        return checkInfoResult;
    }

    /**
     * 针对PPT元素实体校验类对幻灯片对象中具体幻灯片页进行Sound元素校验
     *
     * @param xmlSlideShow            所需校验的幻灯片对象
     * @param elementValidationEntity PPT元素实体校验类
     * @return com.gzmut.office.enums.CheckInfoResult
     */
    public CheckInfoResult checkShape(XMLSlideShow xmlSlideShow, PptElementValidationEntity elementValidationEntity) {
        if (elementValidationEntity.getSlideIndex() > xmlSlideShow.getSlides().size()) {
            return CheckInfoResult.wrong(0, "无法读取第" + (elementValidationEntity.getSlideIndex() + 1) + "张幻灯片");
        }

        XSLFSlide slide = xmlSlideShow.getSlides().get(elementValidationEntity.getSlideIndex());

        Map<String, Object> includeParams = JSON.parseObject(elementValidationEntity.getIncludeParamJson());
        Map<String, Object> excludeParams = JSON.parseObject(elementValidationEntity.getExcludeParamJson());


        CheckInfoResult checkInfoResult = new CheckInfoResult(CheckStatus.CORRECT, 0, System.lineSeparator(), elementValidationEntity.getTargetVerify());
        PptCorrectEnums targetEnum;
        float averageScore = elementValidationEntity.getScore() / (float) (includeParams.size() + excludeParams.size());

        System.out.println("开始" + ParamMatchEnum.INCLUDE.getDesc() + "校验");

        try {
            for (String target : includeParams.keySet()) {
                targetEnum = Objects.requireNonNull(EnumUtils.getEnumByName(PptCorrectEnums.class, target));
                switch (targetEnum) {
                    case NAME: {
                        List<String> correctShapeName = JSON.parseArray(String.valueOf(includeParams.get(target)), ShapeView.class)
                                .stream()
                                .map(ShapeView::getName)
                                .collect(Collectors.toList());
                        List<String> shapeNameParam = pptUtils.getShape(slide)
                                .stream()
                                .map(ShapeView::getName)
                                .collect(Collectors.toList());
                        checkInfoResult = compareSingleLayerParam(checkInfoResult, correctShapeName, shapeNameParam, averageScore, targetEnum.getDesc(), ParamMatchEnum.INCLUDE);
                        break;
                    }
                    case ATTRIBUTE: {
                        checkInfoResult = compareMultiLayerParam(checkInfoResult, includeParams.get(target), pptUtils.getShape(slide), averageScore, targetEnum.getDesc(), ParamMatchEnum.INCLUDE);
                        break;
                    }
                    default:
                        break;
                }
            }
            System.out.println("校验完成");

            System.out.println("开始" + ParamMatchEnum.EXCLUDE.getDesc() + "校验");
            for (String target : excludeParams.keySet()) {
                targetEnum = Objects.requireNonNull(EnumUtils.getEnumByName(PptCorrectEnums.class, target));
                switch (targetEnum) {
                    case NAME: {
                        List<String> correctShapeName = JSON.parseArray(String.valueOf(includeParams.get(target)), ShapeView.class)
                                .stream()
                                .map(ShapeView::getName)
                                .collect(Collectors.toList());
                        List<String> shapeNameParam = pptUtils.getShape(slide)
                                .stream()
                                .map(ShapeView::getName)
                                .collect(Collectors.toList());
                        checkInfoResult = compareSingleLayerParam(checkInfoResult, correctShapeName, shapeNameParam, averageScore, targetEnum.getDesc(), ParamMatchEnum.EXCLUDE);
                        break;
                    }
                    case ATTRIBUTE: {
                        checkInfoResult = compareMultiLayerParam(checkInfoResult, excludeParams.get(target), pptUtils.getShape(slide), averageScore, targetEnum.getDesc(), ParamMatchEnum.EXCLUDE);
                        break;
                    }
                    default:
                        break;
                }
            }
            System.out.println("校验完成");
        } catch (Exception e) {
            e.printStackTrace();
            return CheckInfoResult.exception(0, elementValidationEntity + "解析发生异常:" + e.getMessage());
        }
        return checkInfoResult;
    }

    /**
     * List<String>参数对比（分为匹配模式和排除模式）
     *
     * @param correctSampleList 标准参数对象
     * @param params            所需校验的参数对象
     * @param score             该题总分
     * @param desc              校验描述信息
     * @param matchModel        对比模式（匹配模式/排除模式）
     * @return com.gzmut.office.enums.CheckInfoResult
     */
    public CheckInfoResult compareSingleLayerParam(CheckInfoResult checkInfoResult, List<String> correctSampleList, List<String> params, float score, String desc, ParamMatchEnum matchModel) {
        float averageScore = (correctSampleList.size() == 0) ? score : score / (float) (correctSampleList.size());
        for (String str : correctSampleList) {
            if (matchModel == ParamMatchEnum.INCLUDE) {
                if (correctSampleList.size() == 0 && params.size() != 0 || !params.contains(str)) {
                    checkInfoResult.appendWrongInfoResult(0, desc);
                } else {
                    checkInfoResult.appendCorrectInfoResult(averageScore, desc);
                }
            } else if (matchModel == ParamMatchEnum.EXCLUDE) {
                if (correctSampleList.size() == 0 && params.size() != 0 || !params.contains(str)) {
                    checkInfoResult.appendCorrectInfoResult(averageScore, desc);
                } else {
                    checkInfoResult.appendWrongInfoResult(0, desc);
                }
            }
        }
        return checkInfoResult;
    }

    /**
     * Map<String,String>参数对比（分为匹配模式和排除模式）
     *
     * @param correctSampleObject 标准参数对象
     * @param params              所需校验的参数对象
     * @param score               该题总分
     * @param desc                校验描述信息
     * @param matchModel          对比模式（匹配模式/排除模式）
     * @return com.gzmut.office.enums.CheckInfoResult
     */
    public CheckInfoResult compareSingleLayerParam(CheckInfoResult checkInfoResult, Object correctSampleObject, Map<String, String> params, float score, String desc, ParamMatchEnum matchModel) {
        Map<String, Object> correctSample = JSON.parseObject(String.valueOf(correctSampleObject));
        float averageScore = (correctSample.size() == 0) ? score : score / (float) (correctSample.size());
        for (String key : correctSample.keySet()) {
            if (matchModel == ParamMatchEnum.INCLUDE) {
                if (correctSample.size() == 0 && params.size() != 0 || !correctSample.get(key).equals(params.get(key))) {
                    checkInfoResult.appendWrongInfoResult(0, desc);
                } else {
                    checkInfoResult.appendCorrectInfoResult(averageScore, desc);
                }
            } else if (matchModel == ParamMatchEnum.EXCLUDE) {
                if (correctSample.size() == 0 && params.size() != 0 || !correctSample.get(key).equals(params.get(key))) {
                    checkInfoResult.appendCorrectInfoResult(averageScore, desc);
                } else {
                    checkInfoResult.appendWrongInfoResult(0, desc);
                }
            }
        }
        return checkInfoResult;
    }

    /**
     * JSON多层参数对比（分为匹配模式和排除模式）
     *
     * @param correctSampleObject 标准参数对象
     * @param params              所需校验的参数对象
     * @param score               该题总分
     * @param desc                校验描述信息
     * @param matchModel          对比模式（匹配模式/排除模式）
     * @return com.gzmut.office.enums.CheckInfoResult
     */
    public CheckInfoResult compareMultiLayerParam(CheckInfoResult checkInfoResult, Object correctSampleObject, Object params, float score, String desc, ParamMatchEnum matchModel) {
        JSONArray correctSampleJsonArray = JSON.parseArray(JSON.toJSONString(correctSampleObject));
        JSONArray paramJsonArray = JSON.parseArray(JSON.toJSONString(params));
        float averageScore = (correctSampleJsonArray.size() == 0) ? score : score / (float) (correctSampleJsonArray.size());

        if (matchModel == ParamMatchEnum.INCLUDE) {
            if (correctSampleJsonArray.size() != 0 && paramJsonArray.size() == 0 || correctSampleJsonArray.size() == 0 && paramJsonArray.size() != 0) {
                checkInfoResult.appendWrongInfoResult(0, desc);
            } else {
                for (Object object : correctSampleJsonArray) {
                    JSONObject jsonObject = JSON.parseObject(JSON.toJSONString(object));
                    if (paramJsonArray.contains(jsonObject)) {
                        checkInfoResult.appendCorrectInfoResult(averageScore, desc);
                    } else {
                        checkInfoResult.appendWrongInfoResult(0, desc);
                    }
                }
            }
        } else if (matchModel == ParamMatchEnum.EXCLUDE) {
            if (correctSampleJsonArray.size() != 0 && paramJsonArray.size() == 0 || correctSampleJsonArray.size() == 0 && paramJsonArray.size() != 0) {
                return checkInfoResult.appendCorrectInfoResult(averageScore, desc);
            } else {
                for (Object object : correctSampleJsonArray) {
                    JSONObject jsonObject = JSON.parseObject(JSON.toJSONString(object));
                    if (paramJsonArray.contains(jsonObject)) {
                        checkInfoResult.appendWrongInfoResult(0, desc);
                    } else {
                        checkInfoResult.appendCorrectInfoResult(averageScore, desc);
                    }
                }
            }
        }
        return checkInfoResult;
    }

    /**
     * 文本内容对比（分为匹配模式和排除模式）
     *
     * @param correctSampleObject 标准文本对象
     * @param textContent         所需校验的文本对象
     * @param score               该题总分
     * @param desc                校验描述信息
     * @param matchModel          对比模式（匹配模式/排除模式）
     * @return com.gzmut.office.enums.CheckInfoResult
     */
    public CheckInfoResult compareTextContent(CheckInfoResult checkInfoResult, Object correctSampleObject, String textContent, float score, String desc, ParamMatchEnum matchModel) {
        if (matchModel == ParamMatchEnum.INCLUDE) {
            if (correctSampleObject == null && textContent != null || !String.valueOf(correctSampleObject).equals(textContent)) {
                return checkInfoResult.appendWrongInfoResult(0, desc);
            } else {
                return checkInfoResult.appendCorrectInfoResult(score, desc);
            }
        } else if (matchModel == ParamMatchEnum.EXCLUDE) {
            if (correctSampleObject == null && textContent != null || !String.valueOf(correctSampleObject).equals(textContent)) {
                return checkInfoResult.appendCorrectInfoResult(score, desc);
            } else {
                return checkInfoResult.appendWrongInfoResult(0, desc);
            }
        }
        return checkInfoResult;
    }

    public static String readJsonData(String pactFile) throws IOException {
        StringBuilder stringBuilder = new StringBuilder();
        File myFile = new File(pactFile);//"D:"+File.separatorChar+"DStores.json"
        if (!myFile.exists()) {
            System.err.println("Can't Find " + pactFile);
        }
        try {
            FileInputStream fis = new FileInputStream(pactFile);
            InputStreamReader inputStreamReader = new InputStreamReader(fis, StandardCharsets.UTF_8);
            BufferedReader in = new BufferedReader(inputStreamReader);

            String str;
            while ((str = in.readLine()) != null) {
                stringBuilder.append(str);  //new String(str,"UTF-8")
            }
            in.close();
        } catch (IOException e) {
            e.getStackTrace();
        }
        return stringBuilder.toString();
    }
}