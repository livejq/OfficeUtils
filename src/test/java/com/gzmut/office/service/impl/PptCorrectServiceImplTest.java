package com.gzmut.office.service.impl;

import com.alibaba.fastjson.JSON;
import com.gzmut.office.bean.PptElementValidationEntity;
import com.gzmut.office.bean.PptValidationEntity;
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
    private String includeParamJson = resourcePath + "includeParamJson.json";


    /**
     * JUnit test target exclude param JSON file path
     */
    private String excludeParamJson = resourcePath + "excludeParamJson.json";

    /**
     * JUnit test target PowerPoint file path
     */
    private String fileName = resourcePath + "PPT.pptx";

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
        PptElementValidationEntity pptElementValidationEntity = new PptElementValidationEntity("01", 3.0f, PptTargetEnums.SMART_ART, readJsonData(includeParamJson), readJsonData(excludeParamJson), 3);
        pptElementValidationEntities.add(pptElementValidationEntity);

        PptValidationEntity pptValidationEntity = new PptValidationEntity();
        pptValidationEntity.setId("01")
                .setFileName("PPT.pptx")
                .setPptElementValidationEntities(pptElementValidationEntities);

        long endTime = System.currentTimeMillis();
        System.out.println("初始化判题规则完成——完成时间：" + dateFormat.format(endTime) + "——耗时：" + (endTime - startTime) + "ms");
        checkInfoResults = initCheck(fileName, pptValidationEntity);
        checkInfoResults.forEach(
                checkInfoResult -> System.out.println(System.lineSeparator() + "当前文件——" + fileName + "——校验结果：" + checkInfoResult.getStatus().getDesc() + "——得分：" + checkInfoResult.getScore() + "——详细信息：" + checkInfoResult.getMessage())
        );
    }

    /**
     * PPT文件校验方法（精度：单张幻灯片元素级）
     *
     * @param elementValidationEntity 所需校验的元素实体类
     * @return com.gzmut.office.enums.CheckInfoResult
     */
    public CheckInfoResult check(PptElementValidationEntity elementValidationEntity) {
        if (pptUtils.getSlideShow() == null) return CheckInfoResult.exception(0, "找不到考生文件");

        PptTargetEnums target = elementValidationEntity.getTargetVerify();
        System.out.println("正在校验" + target.getDesc() + "元素");
        switch (target) {
            case SMART_ART:
                return checkSmartArt(pptUtils.getSlideShow(), elementValidationEntity);
            case VIDEO:
                System.out.println("i am come in " + target);
                return null;
            case SOUND:
                System.out.println(" i am come in " + target);
                return null;
            case SHAPE:
                System.out.println("  i am come in " + target);
                return null;
            default:
                return CheckInfoResult.exception(0, "解析目标不存在");
        }
    }

    /**
     * 针对PPT文件校验实体类对其中元素实体类进行遍历校验（精度：PPT文件级）
     *
     * @param filePath 所需检测考生文件路径
     * @param pptValidationEntity PPT校验实体类
     * @return java.util.List<com.gzmut.office.enums.CheckInfoResult>
     */
    public List<CheckInfoResult> initCheck(String filePath, PptValidationEntity pptValidationEntity) {
        long startTime = System.currentTimeMillis();
        System.out.println("开始扫描考生文件——启动时间：" + dateFormat.format(startTime));
        List<CheckInfoResult> checkInfoResults = new ArrayList<>();
        if (!pptUtils.initXMLSlideShow(filePath)) {
            checkInfoResults.add(CheckInfoResult.exception(0, "找不到考生文件"));
        } else {
            List<PptElementValidationEntity> pptElementValidationEntities = pptValidationEntity.getPptElementValidationEntities();
            pptElementValidationEntities.forEach(elementValidationEntity -> checkInfoResults.add(check(elementValidationEntity)));
        }
        long endTime = System.currentTimeMillis();
        System.out.println("扫描完成——完成时间：" + dateFormat.format(endTime) + "——耗时：" + (endTime - startTime) + "ms");
        return checkInfoResults;
    }

    /**
     * 针对PPT元素实体校验类对幻灯片对象中具体幻灯片页进行SmartArt元素校验
     *
     * @param xmlSlideShow 所需校验的幻灯片对象
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
        CheckInfoResult checkInfoResult = new CheckInfoResult(CheckStatus.CORRECT, 0, System.lineSeparator());
        CheckInfoResult tempResult;
        PptCorrectEnums targetEnum;
        float averageScore = elementValidationEntity.getScore() / (float) (includeParams.size() + excludeParams.size());
        try {

            System.out.println("开始"+ParamMatchEnum.INCLUDE.getDesc()+"校验");
            for (String target : includeParams.keySet()) {
                targetEnum = Objects.requireNonNull(EnumUtils.getEnumByName(PptCorrectEnums.class, target));
                switch (targetEnum) {
                    case TEXT_CONTENT:
                        tempResult = compareTextContent(includeParams.get(target), pptUtils.getSmartArtSpliceText(slide), averageScore, targetEnum.getDesc(), ParamMatchEnum.INCLUDE);
                        break;
                    case FORMAT:
                        tempResult = compareParam(includeParams.get(target), pptUtils.getSmartArtLayoutAttributes(slide, PowerPointConstants.SMART_ART_FORMAT_RESOURCE_URL_PREFIX), averageScore, targetEnum.getDesc(),ParamMatchEnum.INCLUDE);
                        break;
                    case COLOR:
                        tempResult = compareParam(includeParams.get(target), pptUtils.getSmartArtLayoutAttributes(slide, PowerPointConstants.SMART_ART_COLOR_RESOURCE_URL_PREFIX), averageScore, targetEnum.getDesc(),ParamMatchEnum.INCLUDE);
                        break;
                    case STYLE:
                        tempResult = compareParam(includeParams.get(target), pptUtils.getSmartArtLayoutAttributes(slide, PowerPointConstants.SMART_ART_STYLE_RESOURCE_URL_PREFIX), averageScore, targetEnum.getDesc(),ParamMatchEnum.INCLUDE);
                        break;
                    default:
                        continue;
                }
                if (tempResult != null) {
                    if (checkInfoResult.getStatus() == CheckStatus.CORRECT && tempResult.getStatus() != CheckStatus.CORRECT) {
                        checkInfoResult.setStatus(tempResult.getStatus());
                    }
                    checkInfoResult.setMessage(checkInfoResult.getMessage() + tempResult.getMessage());
                    checkInfoResult.setScore(checkInfoResult.getScore() + tempResult.getScore());
                }
            }
            System.out.println("校验完成");

            System.out.println("开始"+ParamMatchEnum.EXCLUDE.getDesc()+"校验");
            for (String target : excludeParams.keySet()) {
                targetEnum = Objects.requireNonNull(EnumUtils.getEnumByName(PptCorrectEnums.class, target));
                switch (targetEnum) {
                    case TEXT_CONTENT:
                        tempResult = compareTextContent(excludeParams.get(target), pptUtils.getSmartArtSpliceText(slide), averageScore, targetEnum.getDesc(), ParamMatchEnum.EXCLUDE);
                        break;
                    case FORMAT:
                        tempResult = compareParam(excludeParams.get(target), pptUtils.getSmartArtLayoutAttributes(slide, PowerPointConstants.SMART_ART_FORMAT_RESOURCE_URL_PREFIX), averageScore, targetEnum.getDesc(),ParamMatchEnum.EXCLUDE);
                        break;
                    case COLOR:
                        tempResult = compareParam(excludeParams.get(target), pptUtils.getSmartArtLayoutAttributes(slide, PowerPointConstants.SMART_ART_COLOR_RESOURCE_URL_PREFIX), averageScore, targetEnum.getDesc(),ParamMatchEnum.EXCLUDE);
                        break;
                    case STYLE:
                        tempResult = compareParam(excludeParams.get(target), pptUtils.getSmartArtLayoutAttributes(slide, PowerPointConstants.SMART_ART_STYLE_RESOURCE_URL_PREFIX), averageScore, targetEnum.getDesc(),ParamMatchEnum.EXCLUDE);
                        break;
                    default:
                        continue;
                }
                if (tempResult != null) {
                    if (checkInfoResult.getStatus() == CheckStatus.CORRECT && tempResult.getStatus() != CheckStatus.CORRECT) {
                        checkInfoResult.setStatus(tempResult.getStatus());
                    }
                    checkInfoResult.setMessage(checkInfoResult.getMessage() + tempResult.getMessage());
                    checkInfoResult.setScore(checkInfoResult.getScore() + tempResult.getScore());
                }
            }
            System.out.println("校验完成");
        } catch (Exception e) {
            e.printStackTrace();
            return CheckInfoResult.exception(0, "解析发生异常:" + e.getMessage());
        }
        return checkInfoResult;
    }

    /**
     * Map<String,String>参数对比（分为匹配模式和排除模式）
     *
     * @param correctSampleObject 标准参数对象
     * @param params 所需校验的参数对象
     * @param score 该题总分
     * @param desc 校验描述信息
     * @param matchModel 对比模式（匹配模式/排除模式）
     * @return com.gzmut.office.enums.CheckInfoResult
     */
    public CheckInfoResult compareParam(Object correctSampleObject, Map<String, String> params, float score, String desc, ParamMatchEnum matchModel) {
        Map<String, Object> correctSample = JSON.parseObject(String.valueOf(correctSampleObject));
        CheckInfoResult checkInfoResult = new CheckInfoResult();
        checkInfoResult.setStatus(CheckStatus.CORRECT);
        StringBuilder stringBuilder = new StringBuilder();
        int wrongCount = 0;
        for (String key : correctSample.keySet()) {
            if (matchModel == ParamMatchEnum.INCLUDE) {
                if (!correctSample.get(key).equals(params.get(key))) {
                    wrongCount++;
                    stringBuilder.append(desc).append("——错误×——参数:").append(key).append("——参数值:").append(params.get(key)).append(System.lineSeparator());
                    checkInfoResult.setStatus(CheckStatus.WRONG);
                } else {
                    stringBuilder.append(desc).append("——正确√——参数:").append(key).append("——参数值:").append(params.get(key)).append(System.lineSeparator());
                }
            }else if(matchModel == ParamMatchEnum.EXCLUDE){
                if (!correctSample.get(key).equals(params.get(key))) {
                    stringBuilder.append(desc).append("——正确√——参数:").append(key).append("——参数值:").append(params.get(key)).append(System.lineSeparator());
                } else {
                    wrongCount++;
                    stringBuilder.append(desc).append("——错误×——参数:").append(key).append("——参数值:").append(params.get(key)).append(System.lineSeparator());
                    checkInfoResult.setStatus(CheckStatus.WRONG);
                }
            }
        }
        checkInfoResult.setScore(score - (score * ((float) wrongCount / (float) correctSample.size())));
        checkInfoResult.setMessage(stringBuilder.toString());
        return checkInfoResult;
    }

    /**
     * 文本内容对比（分为匹配模式和排除模式）
     *
     * @param correctSampleObject 标准文本对象
     * @param textContent 所需校验的文本对象
     * @param score 该题总分
     * @param desc 校验描述信息
     * @param matchModel 对比模式（匹配模式/排除模式）
     * @return com.gzmut.office.enums.CheckInfoResult
     */
    public CheckInfoResult compareTextContent(Object correctSampleObject, String textContent, float score, String desc, ParamMatchEnum matchModel) {
        if (matchModel == ParamMatchEnum.INCLUDE) {
            if (!String.valueOf(correctSampleObject).equals(textContent)) {
                return new CheckInfoResult(CheckStatus.WRONG, 0, desc + "——错误×——文本内容:" + textContent+System.lineSeparator());
            } else {
                return new CheckInfoResult(CheckStatus.CORRECT, score, desc + "——正确√——文本内容:" + textContent+System.lineSeparator());
            }
        } else if (matchModel == ParamMatchEnum.EXCLUDE) {
            if (!String.valueOf(correctSampleObject).equals(textContent)) {
                return new CheckInfoResult(CheckStatus.CORRECT, score, desc + "——正确√——文本内容:" + textContent+System.lineSeparator());
            } else {
                return new CheckInfoResult(CheckStatus.WRONG, 0, desc + "——错误×——文本内容:" + textContent+System.lineSeparator());
            }
        }
        return null;
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