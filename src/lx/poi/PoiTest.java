package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.function.Consumer;

import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyles;



public class PoiTest {

    private static XSSFCellStyle cellStyle;

    /**
     *
     * @param args
     * @throws IOException
     */
    public static void main(String[] args) throws IOException {
        // TODO Auto-generated method stub
        //读取word文档流
//		FileInputStream fisSource = new FileInputStream(new File("E:\\Desktop\\source.docx"));
//		//创建XWPFDocument对象
//		XWPFDocument docSource = new XWPFDocument(fisSource);
//		List<XWPFParagraph> paragraphs = docSource.getParagraphs();
//		//不过滤空段，直接遍历
//
//		for(XWPFParagraph p:paragraphs) {
//			List<XWPFRun> runs = p.getRuns();
//			for(XWPFRun r:runs) {
//				//获取run节点的文本
//				String text = r.getText(0);
//				if(text!=null&&text.equals("本科生培养")) {
//					String fontName = r.getFontName();
//					//判断字段
//					if(fontName!=null&&fontName.equals("仿宋")){
//						System.out.println("True");
//					}
//				}
//			}
//		}
//		CTBody body = document.getBody();
//		System.out.println(body);
//		CTSectPr sectPr = document.getDocument().getBody().getSectPr();
//		System.out.println(sectPr);
//		XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
//		int pages= doc.getProperties().getExtendedProperties().getPages();
//		System.out.println(pages);
//		List<IBodyElement> iBody=doc.getBodyElements();
//		List<CTP> pList = body.getPList();
//		System.out.println(pList.size());
        //读取excel文件流
        long start = System.currentTimeMillis();
        FileInputStream fisSource = new FileInputStream(new File("F:\\excel.xlsx"));
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fisSource);
        //获取第一张表
        XSSFSheet sheetAt = xssfWorkbook.getSheetAt(0);
//		int physicalNumberOfRows = sheetAt.getPhysicalNumberOfRows();
//		System.out.println(physicalNumberOfRows);
        //创建一个单元格地址转化
        CellAddress cellAddress = new CellAddress("A1");
        XSSFRow row = sheetAt.getRow(cellAddress.getRow());
        XSSFCell cell = row.getCell(cellAddress.getColumn());
        //获取合并单元格区域
        List<CellRangeAddress> mergedRegions = sheetAt.getMergedRegions();
        for(CellRangeAddress merge :mergedRegions){
            System.out.println(merge);
        }
        //获取单元格类型
        CellType cellType = cell.getCellType();
        //System.out.println(cellType);
        //取值
        String stringCellValue = cell.getStringCellValue();
        //System.out.println(stringCellValue);
        //取背景色
        //字体格式
        PoiTest.cellStyle = cell.getCellStyle();
        XSSFFont font = PoiTest.cellStyle.getFont();
        //下划线0-无，1-有
        byte underline = font.getUnderline();
        //System.out.println(underline);
        //水平对齐
        HorizontalAlignment alignment = cellStyle.getAlignment();
        long end = System.currentTimeMillis();
        System.out.println(end-start);
        //short fillBackgroundColor = PoiTest.cellStyle.getFillBackgroundColor();
        //System.out.println(fillBackgroundColor);
        //XSSFColor fillBackgroundColorColor = PoiTest.cellStyle.getFillBackgroundColorColor();
        //System.out.println(fillBackgroundColorColor.getRGB());
        //System.out.println(fillBackgroundColorColor.getRGB());
        List<String> al = Arrays.asList("a", "b", "c", "d");
        String a ="println";
        //Consumer<String> stringConsumer = System.out::a;
        //al.forEach(stringConsumer);

    }
}
