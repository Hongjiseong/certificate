package service.Impl;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import poi.excel.*;
import service.Certification;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class TableCertification implements Certification {
    public Workbook test(){
        Workbook wb = new SXSSFWorkbook(10000);
        CustomSheet sheet = new CustomSheet(wb);

        CellStyle centerStyle = wb.createCellStyle();
        centerStyle.setAlignment(HorizontalAlignment.CENTER);
        centerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        sheet.setColumnWidth(0, 1500)
             .setColumnWidth(1, 4500)
             .setColumnWidth(2, 4500)
             .setColumnWidth(3, 2000)
             .setColumnWidth(4, 2000)
             .setColumnWidth(5, 4500)
             .setColumnWidth(6, 2000)
             .setColumnWidth(7, 15000)
             .addMergedRegion(0,1,0,1)  // 테이블정의서
             .addMergedRegion(0,0,3,7)  // Database - value
             .addMergedRegion(1,1,3,5)  // Schema
             .addMergedRegion(2,2,0,1)  // 테이블명
             .addMergedRegion(2,2,2,7)  // 테이블명 - value
             .addMergedRegion(3,3,0,1)  // COMMENT
             .addMergedRegion(3,3,2,7); // COMMENT - value


        CustomRow row0 = sheet.createRow(0);
        row0.createCell(0).setCellValue("테이블정의서").setCellStyle(centerStyle);
        row0.createCell(2).setCellValue("Database").setCellStyle(centerStyle);
        row0.createCell(3).setCellValue("   " + "[Database 값 입력]");

        CustomRow row1 = sheet.createRow(1);
        row1.createCell(2).setCellValue("Schema").setCellStyle(centerStyle);
        row1.createCell(3).setCellValue("   " + "[Schema 값 입력]");
        row1.createCell(6).setCellValue("생성일").setCellStyle(centerStyle);
        row1.createCell(7).setCellValue("   " + "[생성일 값 입력]");

        CustomRow row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("테이블명").setCellStyle(centerStyle);
        row2.createCell(2).setCellValue("   " + "[테이블명 값 입력]");

        CustomRow row3 = sheet.createRow(3);
        row3.createCell(0).setCellValue("COMMENT").setCellStyle(centerStyle);
        row3.createCell(2).setCellValue("   " + "[COMMENT 값 입력]");

        CustomRow row4 = sheet.createRow(4);
        row4.createCell(0).setCellValue("Col #").setCellStyle(centerStyle);
        row4.createCell(1).setCellValue("Column Name").setCellStyle(centerStyle);
        row4.createCell(2).setCellValue("Data Type").setCellStyle(centerStyle);
        row4.createCell(3).setCellValue("Key").setCellStyle(centerStyle);
        row4.createCell(4).setCellValue("Null?").setCellStyle(centerStyle);
        row4.createCell(5).setCellValue("Indentity").setCellStyle(centerStyle);
        row4.createCell(6).setCellValue("Default").setCellStyle(centerStyle);
        row4.createCell(7).setCellValue("   " + "Comments");

        List list = new ArrayList();
        for (int i=1; i<=list.size(); i++){
            sheet.createRow(4 + i);
        }

        return wb;
    }

    @Override
    public void download(String path) {
        // excel 파일 저장
        try {
            File xlsFile = new File(path);
            FileOutputStream fileOut = new FileOutputStream(xlsFile);
            test().write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
