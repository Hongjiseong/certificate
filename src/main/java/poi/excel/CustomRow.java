package poi.excel;

import org.apache.poi.ss.usermodel.Row;

public class CustomRow {
    private final Row row;

    public CustomRow(Row row){
        this.row = row;
    }

    public CustomCell createCell(int index){
        return new CustomCell(row.createCell(index));
    }
}
