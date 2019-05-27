package poi.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class CustomSheet {
    private final Sheet sheet;

    public CustomSheet (Workbook wb){
        this.sheet = wb.createSheet("테이블 명세서");
    }

    public CustomSheet setColumnWidth(int index, int width){
        sheet.setColumnWidth(index, width);
        return this;
    }

    public CustomSheet addMergedRegion(int firstRow, int lastRow, int firstCol, int lastCol){
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        return this;
    }

    public CustomRow createRow (int index){
        return new CustomRow(sheet.createRow(index));
    }
}
