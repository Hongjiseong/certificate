package poi.excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class ExcelManager {
    public static Workbook createSXSSF(){
        Workbook wb = new SXSSFWorkbook(10000);
        return wb;
    }
}
