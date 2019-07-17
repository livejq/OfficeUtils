/**
 * @author MXDC
 * @date 2019/7/10
 **/
import com.gzmut.office.*;
import com.gzmut.office.util.*;
public class MainTest {

    public void test(){

    }
    public static void main(String[] args){
        WordUtils wu = new WordUtils();

        wu.getDocument("F:\\new\\请柬1.docx");

        System.out.println(wu.checkAllText("王选"));
        System.out.println(wu.checkAllText("2013"));
        System.out.println(wu.checkDiffFont());
    }
}
