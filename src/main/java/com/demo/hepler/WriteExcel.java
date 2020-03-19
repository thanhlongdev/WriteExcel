package com.demo.hepler;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public class WriteExcel {

    //Write Data
    public void wirteExcel(List<Object[]> list, String path) throws IOException{
        Workbook workbook = createWorkBook(path);
        Sheet sheet = workbook.createSheet();
        int rowIndex = 0;
        //Write Header
        writeHeader(sheet, rowIndex++, "Số Thứ Tự", "Họ Tên");

        //Write Row
        for (Object[] objects: list){
            //Create Row
            Row row = sheet.createRow(rowIndex);
            //Write data on row
            writeBook(objects, row);
            rowIndex++;
        }
        createOutputFile(workbook, path);
    }

    //Create Workbook
    private Workbook createWorkBook(String excelFile) {
        Workbook workbook = null;
        if (excelFile.endsWith("xls")) {
            workbook = new HSSFWorkbook();
        } else if (excelFile.endsWith("xlsx")) {
            workbook = new XSSFWorkbook();
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }
        return workbook;
    }

    //WriteHeader
    private void writeHeader(Sheet sheet, int rowIndex, String... header) {
        //Apply Style
        CellStyle style = createStyleForHeader(sheet);
        //Create Row
        Row row = sheet.createRow(rowIndex);
        for (int i = 0; i < header.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellStyle(style);
            cell.setCellValue(header[i]);
        }
    }

    // Write data
    private void writeBook(Object[] item, Row row) {
        for (int i = 0; i < item.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(item[i].toString());

            //If value is number add below row
            //cell.setCellStyle(setNumberStyle(row));
        }
    }

    private CellStyle createStyleForHeader(Sheet sheet) {
        // Create font
        Font font = sheet.getWorkbook().createFont();
        font.setFontName("Times New Roman");
        font.setBold(true);
        font.setFontHeightInPoints((short) 14); // font size
        font.setColor(IndexedColors.WHITE.getIndex()); // text color

        // Create CellStyle
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        return cellStyle;
    }

    private CellStyle setNumberStyle(Row row){
        short format = (short)BuiltinFormats.getBuiltinFormat("#,##0");
        //Create CellStyle
        Workbook workbook = row.getSheet().getWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(format);
        return cellStyle;
    }

    // Create output file
    private void createOutputFile(Workbook workbook, String excelFilePath) throws IOException {
        OutputStream os = new FileOutputStream(excelFilePath);
        workbook.write(os);
    }

}
