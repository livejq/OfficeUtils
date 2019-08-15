package com.gzmut.office.service.impl;

import com.gzmut.office.bean.CorrectInfo;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;


public class PPTCorrectServiceImplTest {

    @Test
    public void demo01() {
        List<String> jsonArrayList = new ArrayList<>(8);
        String rule1 = "[{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_LEVEL\",\"param\":{\"TITLE\":\"1\",\"FIRST_TEXT\":\"1\",\"SECOND_TEXT\":\"\"},\"location\":{\"lp\":\"1\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_LEVEL\",\"param\":{\"TITLE\":\"1\",\"FIRST_TEXT\":\"5\",\"SECOND_TEXT\":\"\"},\"location\":{\"lp\":\"2\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_LEVEL\",\"param\":{\"TITLE\":\"1\",\"FIRST_TEXT\":\"3\",\"SECOND_TEXT\":\"3\"},\"location\":{\"lp\":\"3\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_LEVEL\",\"param\":{\"TITLE\":\"1\",\"FIRST_TEXT\":\"3\",\"SECOND_TEXT\":\"2\"},\"location\":{\"lp\":\"4\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_LEVEL\",\"param\":{\"TITLE\":\"1\",\"FIRST_TEXT\":\"4\",\"SECOND_TEXT\":\"0\"},\"location\":{\"lp\":\"5\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_LEVEL\",\"param\":{\"TITLE\":\"1\",\"FIRST_TEXT\":\"4\",\"SECOND_TEXT\":\"5\"},\"location\":{\"lp\":\"6\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_LEVEL\",\"param\":{\"TITLE\":\"1\",\"FIRST_TEXT\":\"\",\"SECOND_TEXT\":\"\"},\"location\":{\"lp\":\"8\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_LEVEL\",\"param\":{\"TITLE\":\"1\",\"FIRST_TEXT\":\"\",\"SECOND_TEXT\":\"\"},\"location\":{\"lp\":\"9\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_LEVEL\",\"param\":{\"TITLE\":\"1\",\"FIRST_TEXT\":\"3\",\"SECOND_TEXT\":\"\"},\"location\":{\"lp\":\"13\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_LEVEL\",\"param\":{\"TITLE\":\"1\",\"FIRST_TEXT\":\"\",\"SECOND_TEXT\":\"\"},\"location\":{\"lp\":\"14\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_LEVEL\",\"param\":{\"TITLE\":\"1\",\"FIRST_TEXT\":\"4\",\"SECOND_TEXT\":\"\"},\"location\":{\"lp\":\"15\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_BOX_CONTENT_FORMAT\",\"param\":{\"TEXT_STYLE\":\"\",\"TEXT_SIZE\":\"\",\"TEXT_COLOR\":\"\"},\"location\":{\"lp\":\"1,15\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_SLIDE_NUMBER\",\"param\":{\"COUNT\":\"15\"},\"location\":{\"lp\":\"\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_BOX_CONTENT\",\"param\":{\"TEXT_CONTENT\":\"中国注册税务师协会&#@@#&The china certified tax agents association\"},\"location\":{\"lp\":\"1\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_BOX_CONTENT\",\"param\":{\"TEXT_CONTENT\":\"目录&#@@#&一、协会概况&#@@#&二、协会服务&#@@#&三、教育培训&#@@#&四、党建统战&#@@#&五、联系我们\"},\"location\":{\"lp\":\"2\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_BOX_CONTENT\",\"param\":{\"TEXT_CONTENT\":\"一、协会概况&#@@#&行业概况&#@@#&协会基本情况&#@@#&行业发展经历&#@@#&行业发展规划&#@@#&协会章程&#@@#&组织结构\"},\"location\":{\"lp\":\"3\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_BOX_CONTENT\",\"param\":{\"TEXT_CONTENT\":\"协会基本情况&#@@#&中国注册税务师协会（前身为中国税务咨询协会），成立于1995年2月28日，是经中华人民共和国民政部批准的全国一级社团组织，是由中国注册税务师和税务师事务所组成的行业民间自律管理组织，受民政部和国家税务总局的业务指导和监督管理。英文缩写为CCTAA。协会与多个国际同行业组织建立了友好合作关系，并于2004年11月加入亚洲一大洋洲税务师协会（AOTCA），成为其正式会员。&#@@#&协会宗旨&#@@#&遵守国家法律法规和政策，团结、教育和引导注册税务师及其执业人员，以经济建设为中心，促进社会主义市场经济的发展，维护税收法律法规及规章制度的严肃性，做到正确实施，遵守职业道德，认真履行义务，维护国家利益和会员及被代理人的合法权益，加强会员的联系与合作，协调税务师事务所及其执业人员与委托代理人之间的关系，在政府及有关部门与会员之间起到桥梁和纽带作用，实施注册税务师行业自律管理，促进行业健康发展，为改革开放和繁荣社会主义市场经济服务。&#@@#&协会主要职能&#@@#&“服务、监督、管理、协调”是中国注册税务师协会的主要职能。\"},\"location\":{\"lp\":\"4\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_LAYOUT_STYLE\",\"param\":{\"LAYOUT_STYLE\":\"TITLE_AND_CONTENT\"},\"location\":{\"lp\":\"4\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_BOX_CONTENT\",\"param\":{\"TEXT_CONTENT\":\"行业发展历程&#@@#&至2014年底，全行业有从业人员10万余人，其中执业注册税务师4万多人；税务师事务所5400多家；行业经营收入140多亿元。&#@@#&行业初创：起步于20世纪80年代中期&#@@#&脱钩改制：1999~2000年底&#@@#&规范发展：2005年12月，国家税务总局颁布了《注册税务师管理暂行办法》，又颁发了“十一五”、“十二五”时期中国注册税务师行业发展的指导意见。\"},\"location\":{\"lp\":\"5\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_LAYOUT_STYLE\",\"param\":{\"LAYOUT_STYLE\":\"TITLE_AND_CONTENT\"},\"location\":{\"lp\":\"5\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_BOX_CONTENT\",\"param\":{\"TEXT_CONTENT\":\"行业发展规划&#@@#&2014年，制定《中国注册税务师行业发展规划（2014年——2017年）》。明确行业发展战略、发展目标。&#@@#&发展战略&#@@#&“专业化发展战略”、“规范化发展战略”、“国际化发展战略”、“信息化发展战略”。&#@@#&发展战略&#@@#&“专业化发展战略”、“规范化发展战略”、“国际化发展战略”、“信息化发展战略”。&#@@#&发展目标&#@@#&“把注册税务师行业建设成为具有国际先进水平和中国特色的现代化中介行业。”&#@@#&核心价值观&#@@#&“诚信、公正、敬业、创新”积极倡导诚信服务、公正独立、爱岗敬业、开拓创新；坚决反对弄虚作假、徇私枉法、消极懒散、因循守旧。\"},\"location\":{\"lp\":\"6\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_BOX_CONTENT\",\"param\":{\"TEXT_CONTENT\":\"二、协会服务\"},\"location\":{\"lp\":\"9\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_BOX_CONTENT\",\"param\":{\"TEXT_CONTENT\":\"三、教育培训&#@@#&名师风采&#@@#&面授培训&#@@#&中税协网校\"},\"location\":{\"lp\":\"13\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_LAYOUT_STYLE\",\"param\":{\"LAYOUT_STYLE\":\"TITLE_AND_CONTENT\"},\"location\":{\"lp\":\"13\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_BOX_CONTENT\",\"param\":{\"TEXT_CONTENT\":\"四、党建统战\"},\"location\":{\"lp\":\"14\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_TEXT_BOX_CONTENT\",\"param\":{\"TEXT_CONTENT\":\"五、联系我们&#@@#&地址：北京市海淀区阜成路73号裕惠大厦A座11层&#@@#&邮编：100142&#@@#&电话：010-6841 3988（总机）&#@@#&网址：http://www.cctaa.cn\"},\"location\":{\"lp\":\"15\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"}]";
        String rule2 = "[{\"score\":\"2.0\",\"knowledge\":\"CHECK_SLIDE_THEME\",\"param\":{\"THEME\":\"五彩缤纷\"},\"location\":{\"lp\":\"1,15\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"2.0\",\"knowledge\":\"CHECK_LAYOUT_NAME\",\"param\":{\"LAYOUT_NAME\":\"CUST&#@@#&PIC_TX&#@@#&CUST&#@@#&TEXT&#@@#&VERT_TITLE_AND_TX&#@@#&CUST&#@@#&TITLE_AND_CONTENT&#@@#&BLANK&#@@#&TITLE\"},\"location\":{\"lp\":\"\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"}]";
        String rule3 = "[{\"score\":\"3.0\",\"knowledge\":\"CHECK_LAYOUT_STYLE\",\"param\":{\"LAYOUT_STYLE\":\"TITLE\"},\"location\":{\"lp\":\"1\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"}]";
        String rule4 = "[{\"score\":\"2.0\",\"knowledge\":\"CHECK_TEXT_HYPERLINK\",\"param\":{\"TEXT_HYPERLINK\":\"5\"},\"location\":{\"lp\":\"2\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"2.0\",\"knowledge\":\"CHECK_LAYOUT_STYLE\",\"param\":{\"LAYOUT_STYLE\":\"CUST\"},\"location\":{\"lp\":\"3\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"2.0\",\"knowledge\":\"CHECK_TEXT_ALIGN_STYLE\",\"param\":{\"HORIZONTAL_ALIGN\":\"0,5,0\",\"VERTICAL_ALIGN\":\"0,1,0\"},\"location\":{\"lp\":\"2\",\"ls\":\"\",\"lo\":\"TEXT_BOX\"},\"baseFile\":\"PPT.pptx\"}]";
        String rule5 = "[{\"score\":\"4.0\",\"knowledge\":\"CHECK_LAYOUT_STYLE\",\"param\":{\"LAYOUT_STYLE\":\"TITLE_AND_CONTENT\"},\"location\":{\"lp\":\"8\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"4.0\",\"knowledge\":\"CHECK_LAYOUT_STYLE\",\"param\":{\"LAYOUT_STYLE\":\"BLANK\"},\"location\":{\"lp\":\"7\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"}]";
        String rule6 = "[{\"score\":\"2.0\",\"knowledge\":\"CHECK_LAYOUT_STYLE\",\"param\":{\"LAYOUT_STYLE\":\"CUST\"},\"location\":{\"lp\":\"9\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"2.0\",\"knowledge\":\"CHECK_LAYOUT_STYLE\",\"param\":{\"LAYOUT_STYLE\":\"CUST\"},\"location\":{\"lp\":\"14\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"2.0\",\"knowledge\":\"CHECK_LAYOUT_STYLE\",\"param\":{\"LAYOUT_STYLE\":\"BLANK\"},\"location\":{\"lp\":\"10\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"2.0\",\"knowledge\":\"CHECK_LAYOUT_STYLE\",\"param\":{\"LAYOUT_STYLE\":\"BLANK\"},\"location\":{\"lp\":\"11\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"2.0\",\"knowledge\":\"CHECK_LAYOUT_STYLE\",\"param\":{\"LAYOUT_STYLE\":\"BLANK\"},\"location\":{\"lp\":\"12\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"}]";
        String rule7 = "[{\"score\":\"2.0\",\"knowledge\":\"CHECK_LAYOUT_STYLE\",\"param\":{\"LAYOUT_STYLE\":\"PIC_TX\"},\"location\":{\"lp\":\"15\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"}]";
        String rule8 = "[{\"score\":\"3.0\",\"knowledge\":\"CHECK_SLIDE_NUMBER\",\"param\":{\"COUNT\":\"15\"},\"location\":{\"lp\":\"\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"},{\"score\":\"3.0\",\"knowledge\":\"CHECK_SECTION\",\"param\":{\"SECTION_NAME\":\"默认节&#@@#&概况&#@@#&服务&#@@#&无标题节\",\"SECTION_NUMBER\":\"2,6,4,3\"},\"location\":{\"lp\":\"\",\"ls\":\"\",\"lo\":\"SLIDE\"},\"baseFile\":\"PPT.pptx\"}]";

        jsonArrayList.add(rule1);
        jsonArrayList.add(rule2);
        jsonArrayList.add(rule3);
        jsonArrayList.add(rule4);
        jsonArrayList.add(rule5);
        jsonArrayList.add(rule6);
        jsonArrayList.add(rule7);
        jsonArrayList.add(rule8);

        PPTCorrectServiceImpl pPtCorrect = new PPTCorrectServiceImpl();
        pPtCorrect.setExamDir(".\\temp");
        List<CorrectInfo> correctInfo = pPtCorrect.correct(jsonArrayList);
        for (CorrectInfo info : correctInfo) {
            System.out.println("ppt操作题第 " + info.getNumber() + " 小题, 得分：" + info.getScore() + ", 详细信息：" + System.lineSeparator() + info.getMsg());
        }


        /*for(Map<String, String> map : correctInfo) {
            if(map != null) {
                for (Map.Entry<String, String> entry : map.entrySet()) {
                    String key = entry.getKey();
                    String value = entry.getValue();
                    System.out.println("key =" + key + " value = " + value);
                }
            }
        }*/
    }

    @Test
    public void parseString2xmlTest() {
        String parseStr = null;

        try {
            Document doc = DocumentHelper.parseText("<xml><zhl>zhl</zhl></xml>");
            Element roots = doc.getRootElement();
            System.out.println("根节点 = [" + roots.getName() + "]");
            System.out.println("内容：" + roots.getText());
            parseStr = roots.getText();
            /*//只有根节点……
            Iterator elements=roots.elementIterator();
            while (elements.hasNext()){
                Element child= (Element) elements.next();
                System.out.println("节点名称 = [" + child.getName() + "]"+"节点内容："+child.getText());
                List subElemets=child.elements();
               for(int i=0;i<subElemets.size();i++){
                   Element subChild= (Element) subElemets.get(i);
                   System.out.println("子节点名称："+subChild.getName()+"子节点内容："+subChild.getText());
               }

            }*/

        } catch (DocumentException e) {
            e.printStackTrace();
        }
        System.out.println(parseStr);

    }
}