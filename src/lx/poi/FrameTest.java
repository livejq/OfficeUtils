package poi;

import java.io.File;

public class FrameTest {
    public static void main(String args[]) {
        File file = new File("G:\\计算机设备全年销量统计表");
        int i = new ExcelCorrect().Correct("123", file);
    }
}