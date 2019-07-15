package poi;

import java.io.File;
import java.util.ArrayList;

public class CheckQuestion {
    /**
     * 返回试卷号时使用
     * @param questionNum 试卷号
     * @param filePath 文件路径集合
     */
    public void CheckQuestionnum(String[] filePath, String questionNum)
    {
        File excelfile = new File("");
        File wordfile = new File("");
        File pptfile = new File("");
        int score;//3道操作题成绩
        //textNums = GetTextnum(questionNum);//根据试卷号查找各大题号(152,357,124)
        String textNums = "152,357,124";
        String[] textNum = textNums.split(",");//将大题号拆分
        score =(new ExcelCorrect().Correct(textNum[0],excelfile))+(new WordCorrect().Correct(textNum[1],wordfile))+ (new PPTCorrect().Correct(textNum[2],pptfile));
    }

    /**
     * 返回大题号时使用
     * @param textNum 大题号
     * @param file 文件路径
     */
    public void CheckTextnum(String textNum,File file) throws Exception
    {
        int score;//操作题成绩
        String extend = file.getAbsolutePath();
        if (extend.contains(".xlsx"))
            score = new ExcelCorrect().Correct(textNum, file);
        else if (extend.contains(".docx"))
            score = new WordCorrect().Correct(textNum, file);
        else
            score = new PPTCorrect().Correct(textNum, file);
    }
}
