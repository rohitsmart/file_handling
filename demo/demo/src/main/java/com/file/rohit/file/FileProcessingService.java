package com.file.rohit.file;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

@Service
public class FileProcessingService {

    public void processCsvToExcelAndSave(MultipartFile file, Path filePath) throws IOException {
        String csvContent = new String(file.getBytes(), StandardCharsets.UTF_8);
        String[] rows = csvContent.split("\\r?\\n");
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Formatted Data");
        String[] headers = {
            "Id", "Astrologer Name", "Recruiter", "Name in Bank", 
            "Bank A/C No.", "IFSC Code", "Name in PAN", "PAN No.", "Disable"
        };
        Row headerRow = sheet.createRow(0);
        for (int headerIndex = 0; headerIndex < headers.length; headerIndex++) {
            Cell headerCell = headerRow.createCell(headerIndex);
            headerCell.setCellValue(headers[headerIndex]);
        }
        for (int rowIndex = 0; rowIndex < rows.length; rowIndex++) {
            String[] cells = rows[rowIndex].split(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
            Row excelRow = sheet.createRow(rowIndex + 1);
            for (int cellIndex = 0; cellIndex < cells.length; cellIndex++) {
                Cell cell = excelRow.createCell(cellIndex);
                String cellValue = cells[cellIndex].trim().replaceAll("^\"|\"$", "");
                cellValue = cellValue.replaceAll("^'|'$", "");
                cell.setCellValue(cellValue);
            }
        }
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            workbook.write(outputStream);
            Files.write(filePath, outputStream.toByteArray());
        } finally {
            workbook.close();
        }
    }

}
