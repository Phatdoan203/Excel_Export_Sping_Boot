package com.api.excel.exporter.controller;

import com.api.excel.exporter.service.ExcelExporterService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;


@Slf4j
@RestController
@RequiredArgsConstructor
@RequestMapping("/excel")
public class ExcelExporterController {
    private final ExcelExporterService excelExporterService;
    @Value(value = "${templates.filename}")
    private String template;

    @GetMapping("/export")
    public ResponseEntity<byte[]> exportExcel(
            @RequestParam(value = "template", defaultValue = "rp-template.xlsx") String templateFileName) {
        try {
            // Create a unique filename for the output
            String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
            String outputFileName = "products_export_" + timestamp + ".xlsx";

            // Create a temporary file path
            Path tempPath = Files.createTempFile("excel_export_", ".xlsx");
            String outputPath = tempPath.toString();

            log.info("Starting Excel export with template: {}, output path: {}", templateFileName, outputPath);

            // Call the service to generate the Excel file
            long startTime = System.currentTimeMillis();
            excelExporterService.exportProductsToExcel(templateFileName, outputPath);
            long endTime = System.currentTimeMillis();

            log.info("Excel export completed in {} ms", (endTime - startTime));

            // Read the file back as bytes
            File file = new File(outputPath);
            FileInputStream fileInputStream = new FileInputStream(file);
            byte[] fileBytes = new byte[(int) file.length()];
            fileInputStream.read(fileBytes);
            fileInputStream.close();

            // Clean up the temporary file
            Files.deleteIfExists(tempPath);

            // Configure response headers
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDispositionFormData("attachment", outputFileName);
            headers.setContentLength(fileBytes.length);

            return new ResponseEntity<>(fileBytes, headers, HttpStatus.OK);

        } catch (IOException e) {
            log.error("Error exporting Excel file", e);
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}
