package com.example.excelexporttiny.service;

import org.apache.commons.codec.binary.Base64;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;

@Service
public class ExportService {
    private transient HSSFWorkbook workbook;
    private transient HSSFSheet sheet;


    public ExportService() {
        this.workbook = new HSSFWorkbook();
    }

    private void writeHeaderLine() {
        sheet = workbook.createSheet("Employee Metrics");

        Row row = sheet.createRow(0);

        CellStyle style = workbook.createCellStyle();
        HSSFFont font = workbook.createFont();
        font.setBold(true);
//        font.setFontHeight(16);
//        font.setCharSet(162);
        style.setFont(font);

        createCell(row, 0, "Adı", style);
        createCell(row, 1, "Soyadı", style);
        createCell(row, 2, "Doğum Tarihi", style);
        createCell(row, 3, "Sınıfı ", style);
        createCell(row, 4, "Mesleği ", style);
    }

    private void createCell(Row row, int columnCount, Object value, CellStyle style) {
        DateTimeFormatter inputFormat = DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss");
        sheet.autoSizeColumn(columnCount);
        Cell cell = row.createCell(columnCount);
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof LocalDateTime) {
            cell.setCellValue(((LocalDateTime) value).format(inputFormat));
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else {
            cell.setCellValue((String) value);
        }

        cell.setCellStyle(style);
    }

    public ResponseEntity<String> export(HttpHeaders response) throws IOException {
        writeHeaderLine();


        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        workbook.write(bos);
        workbook.close();
        bos.close();
        byte[] data = bos.toByteArray();
        for(byte e:data){
            System.out.println(e);
        }

        String result = Base64.encodeBase64String(bos.toByteArray());

        return ResponseEntity.ok().headers(response).body(result);
    }

    public byte[] base64ToByteArray(String base64){
        return Base64.decodeBase64(base64);
    }




}
