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

    // Method to process CSV and save it as Excel
    public void processCsvToExcelAndSave(MultipartFile file, Path filePath) throws IOException {
        // Read CSV file content
        String csvContent = new String(file.getBytes(), StandardCharsets.UTF_8);
        
        // Split CSV into rows by newline
        String[] rows = csvContent.split("\\r?\\n"); // Handle different line endings
        
        // Create new Excel workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Formatted Data");

        try {
            // Loop through the CSV rows
            for (int rowIndex = 0; rowIndex < rows.length; rowIndex++) {
                // Split row into individual cells, considering commas inside quotes
                String[] cells = rows[rowIndex].split(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
                
                // Create a new row in the Excel sheet
                Row excelRow = sheet.createRow(rowIndex);

                // Loop through each cell in the row
                for (int cellIndex = 0; cellIndex < cells.length; cellIndex++) {
                    // Create a new cell in the Excel row
                    Cell cell = excelRow.createCell(cellIndex);

                    // Clean up the cell value by trimming and removing single quotes
                    String cellValue = cells[cellIndex].trim().replaceAll("^\"|\"$", ""); // Remove double quotes
                    cellValue = cellValue.replaceAll("^'|'$", ""); // Remove single quotes

                    // Set cleaned value in the Excel cell
                    cell.setCellValue(cellValue);
                }
            }

            // Save the generated Excel file to the specified path
            try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                workbook.write(outputStream);
                Files.write(filePath, outputStream.toByteArray());
            }
        } finally {
            // Close the workbook to avoid memory leaks
            workbook.close();
        }
    }
}
