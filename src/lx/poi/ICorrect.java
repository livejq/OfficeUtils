package poi;

import java.io.File;
//Word/Excel/PPT单道大题批改接口
public interface ICorrect {
    /**
     * 解析JSON，判分
     * @param textNum 大题号，前端提供
     * @param file 要批改的文件
     * @return 返回一道大题的分数
     */
    public abstract int Correct(String textNum, File file);

    /**
     * 根据id选择工具类的方法
     * @param id   工具类的方法
     * @param value   答案
     * @param lp    Excel表名/Word段落/PPT页
     * @param ls    Excel单元格/Word指定字符串/PPT？？？
     * @param file  要批改的文件
     * @return  返回错误信息
     */
    public abstract String useToGetUtils(int id, String value, String lp, String ls, File file);
}
