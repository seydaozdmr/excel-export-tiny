package com.example.excelexporttiny.controller;

import com.example.excelexporttiny.service.ExportService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

@RestController
public class ExportController {

    private final ExportService exportService;

    public ExportController(ExportService exportService) {
        this.exportService = exportService;
    }

    @PostMapping("/export")
    public ResponseEntity<String> exportToExcel() throws IOException {
        HttpHeaders response = new HttpHeaders();
        response.setContentType(new MediaType("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        DateFormat dateFormatter = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
        String currentDateTime = dateFormatter.format(new Date());

        String headerKey = "Content-Disposition";
        String headerValue = "attachment" + "; filename=EmployeeMetrics_" + currentDateTime + ".xlsx";
        response.set(headerKey, headerValue);

        ResponseEntity<String> export= exportService.export(response);

        response.setContentLength(export.getBody().length());

        return export;

    }
}
