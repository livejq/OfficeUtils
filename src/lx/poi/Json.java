package poi;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Json {
    public static void main(String[] args) {

        //System.out.println(getrow("B2")+" "+getcolum("B2"));
        getCellsList("666");
        //ArrayList arrayList = new ArrayList();

/*        JSONArray array = JSON.parseArray(json6);

        //arrayList = JSONObject.parseObject(json, ArrayList.class);
        List<CheckBean> list = JSONObject.parseArray(array.toJSONString(), CheckBean.class);
        int test=0;
        for(CheckBean bean:list){
            System.out.println(test);
            for (String id : bean.param.keySet()) {
                //System.out.println(list.get(0).num+"------"+id+"------"+list.get(0).param.get(id));
                sheetname(Integer.parseInt(id),bean.param.get(id).toString(),bean.location.lp,bean.location.ls);
            }
            test++;
        }*/
        //workBook.getSheet(sheetname).getRow(getcolum(cell)-1).getCell(getrow(cell)-1).
        //writejson();
    }
    public static void sheetname(int id,String value,String sheetname,String cell){
        switch (id){
            case 1:
                System.out.println("case 1:需要检查该Excel是否存在表："+ sheetname);
                break;
            case 2:
                if(value.contains(":"))
                {
                    System.out.println("case 2:需要检查\""+sheetname+"\"表的\""+cell+"\"单元格里的\"内容\"是否为\""+value+"\"");
                }
                else{
                    System.out.println("case 2:需要检查\""+sheetname+"\"表的\""+cell+"\"单元格里的\"内容\"是否为\""+value+"\"");
                }
                break;
            case 3:
                System.out.println("case 3:需要检查\""+sheetname+"\"表的\""+cell+"\"单元格里的\"字号\"是否不为\""+value+"\"");
                break;
            case 4:
                System.out.println("case 4:需要检查\""+sheetname+"\"表的\""+cell+"\"单元格里的\"字号\"是否大于\""+value+"\"");
                break;
            case 5:
                System.out.println("case 5:需要检查\""+sheetname+"\"表的\""+cell+"\"单元格里的\"颜色\"是否不为\""+value+"\"");
                break;
            case 6:
                System.out.println("case 6:需要检查\""+sheetname+"\"表的\""+cell+"\"单元格是否\"合并\"");
                break;
            case 7:
                System.out.println("case 7:需要检查\""+sheetname+"\"表的\""+cell+"\"单元格里的\"对齐方式\"是否为\""+value+"\"");
                break;
            case 8:
                System.out.println("case 8:需要检查\""+sheetname+"\"表的\""+cell+"\"单元格的\"行高\"是否大于\""+value+"\"");
                break;
            case 9:
                System.out.println("case 9:需要检查\""+sheetname+"\"表的\""+cell+"\"单元格的\"列宽\"是否大于\""+value+"\"");
                break;
            case 10:
                System.out.println("case 10:需要检查\""+sheetname+"\"表是否存在\"图表\"");
                break;
            case 11:
                System.out.println("case 11:需要检查\""+sheetname+"\"表的\""+cell+"\"单元格的\"数值格式\"是否为\""+value+"\"");
                break;
            case 12:
                System.out.println("case 12:需要检查\""+sheetname+"\"表的\""+cell+"\"单元格是否存在\"边框线\"");
                break;
            case 13:
                System.out.println("case 13:需要检查\""+sheetname+"\"表的\""+cell+"\"单元格的\"名称\"是否为\""+value+"\"");
                break;
            case 14:
                System.out.println("case 14:需要检查\""+sheetname+"\"表的\""+cell+"\"单元格是否使用\"公式\"");
                break;
            case 15:
                System.out.println("case 15:需要检查\""+sheetname+"\"表是否存在\"数据透视表\"");
                break;
            case 16:
                break;
            case 17:
                break;

        }

    }
    public static void getCellsList(String cellsarea){
        cellsarea="A3:B6";
        String top = cellsarea.substring(0, cellsarea.indexOf(":"));
        String buttom = cellsarea.substring(cellsarea.indexOf(":")+1,cellsarea.length());
        CellRangeAddress cellAddresses = new CellRangeAddress( getrow(top),getrow(buttom),getcolumn(top),getcolumn(buttom));
        List<String> listcells = new ArrayList<String>();
        for(int i=getrow(top);i<=getrow(buttom);i++){
            for(int j = (getcolumn(top)-1);j<=(getcolumn(buttom)-1);j++){
                char letter = (char)((int)'A'+j);
                //System.out.println(letter+""+i);
                listcells.add(letter+""+i);
            }
        }
        for(String s : listcells){
            System.out.println(s);
        }
    }
    public static void writejson(){
        ArrayList<CheckBean> List = new ArrayList<CheckBean>();

        CheckBean checkBean = new CheckBean();
        checkBean.knowledge="15";
        CheckBean.Location location = new CheckBean.Location();
        location.setLp("数据透视分析");//表名
        location.setLs("");//单元格
        checkBean.setLocation(location);
        Map<String,Object> param = new HashMap();
        param.put("15","true");
        checkBean.score = 2;
        checkBean.setParam(param);
        List.add(checkBean);

        /*checkBean = new CheckBean();
        checkBean.knowledge="2";
        location = new CheckBean.Location();
        location.setLp("数据透视分析");
        location.setLs("");
        checkBean.setLocation(location);
        param = new HashMap();
        param.put("15","true");
        checkBean.score = 2;
        checkBean.setParam(param);
        List.add(checkBean);*/

        System.out.println(JSONObject.toJSON(List));

    }
    public static int getcolumn(String cell){//cell="AC5"
        Map<Character,Integer> map = new HashMap();
        map.put('A',1);
        map.put('B',2);
        map.put('C',3);
        map.put('D',4);
        map.put('E',5);
        map.put('F',6);
        map.put('G',7);
        map.put('H',8);
        map.put('I',9);
        map.put('J',10);
        map.put('K',11);
        map.put('L',12);
        map.put('M',13);
        map.put('N',14);
        map.put('O',15);
        map.put('P',16);
        map.put('Q',17);
        map.put('R',18);
        map.put('S',19);
        map.put('T',20);
        map.put('U',21);
        map.put('V',22);
        map.put('W',23);
        map.put('X',24);
        map.put('Y',25);
        map.put('Z',26);
        int row = 0;
        String reg = "[^a-zA-Z]";
        cell = cell.replaceAll(reg,"");
        char[] cell1 = cell.toCharArray();
        for(int i = 0,j=cell1.length-1;i<cell1.length;i++,j--){
            row+=map.get(cell1[i])*Math.pow(26,j);
        }
        return row;
    }
    public static int getrow(String cell){
        String reg = "[!^a-zA-Z]";
        int num =(Integer.parseInt(cell.replaceAll(reg,"")));
        return num;
    }
}
