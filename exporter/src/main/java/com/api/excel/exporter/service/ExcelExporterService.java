package com.api.excel.exporter.service;

import com.api.excel.exporter.entity.Product;
import com.api.excel.exporter.repository.ProductRepository;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

@Slf4j
@RequiredArgsConstructor
@Service
public class ExcelExporterService {
    private final ProductRepository productRepository;

    public void exportProductsToExcel(String templateFileName, String outputPath) throws IOException {
        if (templateFileName == null || templateFileName.isEmpty()) {
            throw new IllegalArgumentException("Template file name must not be null or empty");
        }

        File templateFile = new ClassPathResource("templates/" + templateFileName).getFile();
        if (!templateFile.exists()) {
            throw new FileNotFoundException("Template file not found: " + templateFileName);
        }

        // Đọc toàn bộ template vào bộ nhớ và sử dụng nó trực tiếp
        try (FileInputStream fis = new FileInputStream(templateFile);
             XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis);
             SXSSFWorkbook workbook = new SXSSFWorkbook(xssfWorkbook, 1000);
             ) {

            fis.close(); // Đóng ngay sau khi đọc xong
            workbook.setCompressTempFiles(true);


            SXSSFSheet sheet =  workbook.getSheetAt(0);


//            SXSSFSheet sheet = workbook.getSheetAt(0);
            int startRow = 15;
            int currentRow = startRow;
            int pageSize = 5000; // Giảm kích thước trang để tránh sử dụng quá nhiều bộ nhớ

            // Cache styles để tối ưu hóa hiệu suất
            Map<Integer, CellStyle> columnStyles = new HashMap<>();
//            if (sheet.getRow(startRow - 1) != null) {
//                // Lấy style từ hàng trên cùng của dữ liệu (thường là hàng tiêu đề)
//                Row headerRow = sheet.getRow(startRow - 1);
//                for (Cell cell : headerRow) {
//                    columnStyles.put(cell.getColumnIndex(), cell.getCellStyle());
//                }
//            }

            log.info("Starting Excel export with template: {}", templateFileName);
            long startTime = System.currentTimeMillis();
            int totalProducts = 0;

            for (int pageNum = 0; ; pageNum++) {
                Page<Product> productPage = productRepository.findAll(PageRequest.of(pageNum, pageSize));
                if (productPage.isEmpty()) break;

                for (Product product : productPage.getContent()) {
                    Row row = sheet.createRow(currentRow);

                    createCell(row, 1, product.getId(), columnStyles.get(1));
                    createCell(row, 2, product.getName(), columnStyles.get(2));
                    createCell(row, 3, product.getCategory(), columnStyles.get(3));
                    createCell(row, 4, product.getPrice().doubleValue(), columnStyles.get(4));
                    createCell(row, 5, product.getStockQuantity(), columnStyles.get(5));
                    createCell(row, 6, product.getDateAdded().toString(), columnStyles.get(6));

                    currentRow++;
                }

                totalProducts += productPage.getContent().size();
                log.info("Processed page {} with {} products, current row: {}",
                        pageNum, productPage.getContent().size(), currentRow);

                if (pageNum % 1000 == 0) {
                    sheet.flushRows(1000);  // Xóa 100 hàng đầu tiên khỏi bộ nhớ
                }

                // Đánh dấu để GC thu hồi bộ nhớ
                System.gc();
            }

            // Tự động điều chỉnh chiều rộng cột
            for (int i = 1; i <= 6; i++) {
                sheet.autoSizeColumn(i);
            }

            // Lưu workbook
            try (FileOutputStream outputStream = new FileOutputStream(outputPath)) {
                workbook.write(outputStream);
                workbook.dispose(); // Giải phóng bộ nhớ
                outputStream.close();
                outputStream.flush();
            }

            long endTime = System.currentTimeMillis();
            log.info("Excel export completed in {} ms, total products: {}",
                    (endTime - startTime), totalProducts);
        }
    }

    private void createCell(Row row, int column, Object value, CellStyle style) {
        Cell cell = row.createCell(column);

        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Long) {
            cell.setCellValue((Long) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value != null) {
            cell.setCellValue(value.toString());
        }

        if (style != null) {
            // Directly apply the existing style instead of creating a new one
            cell.setCellStyle(style);
        }
    }
}
