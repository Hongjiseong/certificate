package poi.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

public class CustomCell {
    private final Cell cell;

    public CustomCell(Cell cell){
        this.cell = cell;
    }

    public CustomCell setCellValue(String value){
        cell.setCellValue(value);
        return this;
    }

    public CustomCell setCellStyle(CellStyle cellStyle){
        cell.setCellStyle(cellStyle);
        return this;
    }
}
