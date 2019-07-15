package poi;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

//Excel批改器
public class ExcelCorrect implements ICorrect{

    /**
     * 解析JSON，判分
     * @param textNum 大题号，前端提供
     * @param file 要批改的文件
     * @return 返回一道大题的分数
     */
    @Override
    public int Correct(String textNum, File file) {

        String json1 = "[{\"score\":2,\"param\":{\"1\":\"true\"},\"location\":{\"lp\":\"平均单价\",\"ls\":\"\"},\"knowledge\":\"1\"},{\"score\":2,\"param\":{\"1\":\"true\"},\"location\":{\"lp\":\"销售情况\",\"ls\":\"\"},\"knowledge\":\"1\"}]";
        String json2 = "[{\"score\":4,\"param\":{\"2\":\"序号\"},\"location\":{\"lp\":\"销售情况\",\"ls\":\"A3\"},\"knowledge\":\"2\"},{\"score\":4,\"param\":{\"2\":{\"lp\":\"销售情况\",\"ls\":\"A4:A83\"}},\"location\":{\"lp\":\"销售情况\",\"ls\":\"A4:A83\"},\"knowledge\":\"2\"}]";
        String json3 = "[{\"score\":2,\"param\":{\"3\":\"宋体\",\"4\":\"11\",\"6\":\"true\",\"5\":\"0,0,0\",\"6\":\"true\",\"7\":\"CENTER\"},\"location\":{\"lp\":\"销售情况\",\"ls\":\"B1\"},\"knowledge\":\"2\"},{\"score\":3,\"param\":{\"11\":\"保留两位小数\",\"7\":\"CENTER\",\"12\":\"true\",\"8\":\"14.25\",\"9\":\"8.38\"},\"location\":{\"lp\":\"销售情况\",\"ls\":\"A3:F83\"},\"knowledge\":\"2\"}]";
        String json4 = "[{\"score\":2,\"param\":{\"13\":\"商品均价\"},\"location\":{\"lp\":\"平均单价\",\"ls\":\"B3:C7\"},\"knowledge\":\"2\"},{\"score\":3,\"param\":{\"14\":\"true\",\"2\":{\"lp\":\"销售情况\",\"ls\":\"F\"}},\"location\":{\"lp\":\"销售情况\",\"ls\":\"F\"},\"knowledge\":\"2\"}]";
        String json5 = "[{\"score\":2,\"param\":{\"1\":\"true\"},\"location\":{\"lp\":\"数据透视分析\",\"ls\":\"\"},\"knowledge\":\"1\"}, {\"score\":2,\"param\":{\"15\":\"true\"},\"location\":{\"lp\":\"数据透视分析\",\"ls\":\"\"},\"knowledge\":\"1\"}]";
        String json6 = "[{\"score\":2,\"param\":{\"10\":\"true\"},\"location\":{\"lp\":\"数据透视分析\",\"ls\":\"\"},\"knowledge\":\"15\"}]";

        //String[] json = getjson(textNum);//从数据库获取各小题规则json,在Json类内写静态方法，数组长度就是小题数
        ArrayList<String> jsons = new ArrayList<String>();
        //jsons.add(json1);
        //jsons.add(json2);
        jsons.add(json3);
        jsons.add(json4);
        jsons.add(json5);
        jsons.add(json6);
        /**
         * rulescore规则数
         * truescore正确数
         * score=truescore/rulescore*知识点分数
         */
        int rulescore=0,truescore=0;
        int finalscore=0;//一大题的分数
        double score=0.0;
        JSONArray array;

        for(int i = 0; i < jsons.size(); i++) { //小题数
            array = JSON.parseArray(jsons.get(i));
            List<CheckBean> list = JSONObject.parseArray(array.toJSONString(), CheckBean.class);//将json反序列为List<CheckBean>对象
            truescore=0;
            rulescore=0;
            score=0.0;
            for (CheckBean bean : list) {//知识点

                for (String id : bean.param.keySet()) {//单个属性
                    String result = useToGetUtils(Integer.parseInt(id), bean.param.get(id).toString(), bean.location.lp, bean.location.ls, file);
                    if (result.equals("true")){
                        System.out.println("true");
                        truescore++;
                    }
                    else
                        System.out.println(result+"储存到日志");
                    //WriteToMySql(result);
                    rulescore++;
                }
                score = (truescore / rulescore+0.5) * bean.score;//bean.score每个知识点的成绩
                finalscore+=score;
                System.out.println("正确数："+truescore+",总数："+rulescore+",得分:"+score+",总分："+finalscore);
            }

        }
        return finalscore;
    }

    /**
     * 根据id选择工具类的方法
     * @param id   工具类的方法
     * @param value   答案
     * @param lp    Excel表名/Word段落/PPT页
     * @param ls    Excel单元格/Word指定字符串/PPT？？？
     * @param file  要批改的文件
     * @return  返回错误信息
     */
    @Override
    public String useToGetUtils(int id, String value, String lp, String ls, File file) {
        ExcelUtils eu = new ExcelUtils();
        eu.setWorkbook(file.getAbsolutePath()+".xlsx");
        switch (id) {
            case 1://需要检查该Excel是否存在表
                System.out.println("case 1:需要检查该Excel是否存在表:"+lp+"");
                return eu.checkSheetName(lp) ? "true" : "false:不存在表名为"+lp+"的表";
            case 2:
                if(value.contains(":"))//单元格区域
                {
                    CheckBean.Location value1 = JSONObject.parseObject(value, CheckBean.Location.class);
                    //CheckBean.Location value1 = (CheckBean.Location)value;
                    System.out.println("case 2:需要检查\""+lp+"\"表的\""+ls+"\"单元格里的\"内容\"是否与标准答案文档里的\""+value1.lp+"\"表的\""+value1.ls+"\"单元格里的内容相同");
                    return "true";
                }
                else{//单元格
                    System.out.println("case 2:需要检查\""+lp+"\"表的\""+ls+"\"单元格里的\"内容\"是否为\""+value+"\"");
                    if(eu.getCellValue(lp,ls).equals(value))
                        return  "true";
                    else
                        return "false,单元格内容为:\""+eu.getCellValue(lp,ls)+"\"与答案\""+value+"\"不符合,";
                }
            case 3:
                System.out.println("case 3:需要检查\""+lp+"\"表的\""+ls+"\"单元格里的\"字体\"是否不为\""+value+"\"");
                String fontName = eu.getCellFont(lp,ls);
                if(!fontName.equals(value))
                    return "true";
                else
                    return "false,字体为:\""+fontName+"\"，与默认字体相同\""+value+"\"";
            case 4:
                System.out.println("case 4:需要检查\""+lp+"\"表的\""+ls+"\"单元格里的\"字号\"是否大于\""+value+"\"");
                int i =Integer.parseInt(eu.getCellFontHeight(lp,ls))/20;
                if(i>Integer.parseInt(value))
                    return "true";
                else
                    return "false,字号为:\""+i+"\"不大于答案\""+value+"\",";
            case 5:
                System.out.println("case 5:需要检查\""+lp+"\"表的\""+ls+"\"单元格里的\"颜色\"是否不为\""+value+"\"");
                String[] textNum = value.split(",");
                byte[] bytes =eu.getCellFontColor(lp,ls);
                boolean result1,result2,result3,result;
                result1=textNum[0].equals(((Object)bytes[0]).toString());
                result2=textNum[1].equals(((Object)bytes[1]).toString());
                result3=textNum[2].equals(((Object)bytes[2]).toString());
                result = result1&result2&result3;
                if(!result)
                    return "true";
                else
                    return "false,颜色为:\""+bytes[0]+","+bytes[1]+","+bytes[2]+"\"与为答案相同\""+textNum[0]+","+textNum[1]+","+textNum[2]+"\",";
            case 6:
                System.out.println("case 6:需要检查\""+lp+"\"表的\""+ls+"\"单元格是否\"合并\"");
                if(eu.isCellMerge(lp,ls))
                    return "true";
                else
                    return "false,"+ls+"单元格没有合并，";
            case 7:
                System.out.println("case 7:需要检查\""+lp+"\"表的\""+ls+"\"单元格里的\"对齐方式\"是否为\""+value+"\"");
                if(ls.contains(":")) {//单元格区域
                    List<String> cellsList= eu.getCellsList(ls);
                    int sss =1;
                    for(String cell : cellsList)
                    {
                        String str = eu.getCellAlignment(lp, cell);
                        System.out.println("单元格区域判断对齐方式次数"+(sss++));
                        if(!str.equals(value))
                            return "false,"+cell+"单元格的对齐方式为\"" + str + "\"与答案不符合\"" + value + "\",";
                    }
                    return "true";
                } else {//一个单元格
                    String alignment = eu.getCellAlignment(lp, ls);
                    if (alignment.equals(value))
                        return "true";
                    else
                        return "false," + ls + "的对齐方式为\"" + alignment + "\"与答案不符合\"" + value + "\",";
                }
            case 8:
                System.out.println("case 8:需要检查\""+lp+"\"表的\""+ls+"\"单元格的\"行高\"是否大于\""+value+"\"");
                if(ls.contains(":")) {//单元格区域
                    List<String> cellsList= eu.getCellsList(ls);
                    int sss =1;
                    for(String cell : cellsList)
                    {
                        String str = eu.getCellHeight(lp, cell);
                        System.out.println("单元格区域判断行高次数"+(sss++));
                        if(!str.equals(value))
                            return "false,"+cell+"单元格的行高为\"" + str + "\"与答案不符合\"" + value + "\",";
                    }
                    return "true";
                } else {//一个单元格
                    String height = eu.getCellHeight(lp, ls);
                    if (height.equals(value))
                        return "true";
                    else
                        return "false," + ls + "的行高为\"" + height + "\"与答案不符合\"" + value + "\",";
                }
            case 9:
                System.out.println("case 9:需要检查\""+lp+"\"表的\""+ls+"\"单元格的\"列宽\"是否大于\""+value+"\"");
                return "true";
            case 10:
                System.out.println("case 10:需要检查\""+lp+"\"表是否存在\"簇行图\"");
                return "true";
            case 11:
                System.out.println("case 11:需要检查\""+lp+"\"表的\""+ls+"\"单元格的\"数值格式\"是否为\""+value+"\"");
                return "true";
            case 12:
                System.out.println("case 12:需要检查\""+lp+"\"表的\""+ls+"\"单元格是否存在\"边框线\"");
                return "true";
            case 13:
                System.out.println("case 13:需要检查\""+lp+"\"表的\""+ls+"\"单元格的\"名称\"是否为\""+value+"\"");
                return "true";
            case 14:
                System.out.println("case 14:需要检查\""+lp+"\"表的\""+ls+"\"单元格是否使用\"公式\"");
                return "true";
            case 15:
                System.out.println("case 15:需要检查\""+lp+"\"表是否存在\"数据透视表\"");
                return "true";
            default:
                return "False,Error,发生未知错误";
        }
    }
}
