package com.file.rohit.file;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

@Service
public class FileCompareService {

    public void compareAndUpdateFiles(MultipartFile softFile, MultipartFile hardFile, Path newFilePath) throws IOException {
        // Load the Excel files
        try (Workbook softWorkbook = new XSSFWorkbook(softFile.getInputStream());
             Workbook hardWorkbook = new XSSFWorkbook(hardFile.getInputStream())) {

            Sheet softSheet = softWorkbook.getSheetAt(0); // Assuming data is in the first sheet
            Sheet hardSheet = hardWorkbook.getSheetAt(0); // Assuming data is in the first sheet

            // Create a map to hold the data from the hard manual file (using ID as key)
            Map<String, Row> hardDataMap = new HashMap<>();
            for (Row row : hardSheet) {
                if (row.getRowNum() == 0) continue; // Skip header row

                // Assuming the first column is the ID
                String id = getCellValue(row.getCell(0));  // Get ID as a string
                hardDataMap.put(id, row);
            }

            // Track if new rows are added or updated
            boolean updated = false;

            // Loop through the soft file and compare/update with the hard file data
            for (Row row : softSheet) {
                if (row.getRowNum() == 0) continue; // Skip header row

                String id = getCellValue(row.getCell(0)); // Get ID as a string
                Row hardRow = hardDataMap.get(id);

                if (hardRow == null) {
                    // ID does not exist in the hard file, add a new row
                    Row newRow = hardSheet.createRow(hardSheet.getPhysicalNumberOfRows());
                    for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
                        Cell softCell = row.getCell(i);
                        Cell newCell = newRow.createCell(i);
                        newCell.setCellValue(getCellValue(softCell)); // Copy the soft file data to the new row
                    }
                    updated = true; // Mark as updated
                } else {
                    // Compare the columns (e.g., Name, PAN No., Account details)
                    boolean dataChanged = false;
                    for (int i = 1; i <= 7; i++) { // Assuming columns 1 to 7 are the data columns to compare
                        String softData = getCellValue(row.getCell(i));
                        String hardData = getCellValue(hardRow.getCell(i));

                        if (!softData.equals(hardData)) {
                            // If data is different, update the hard file row with the soft file data
                            hardRow.getCell(i).setCellValue(softData);
                            dataChanged = true;
                        }
                    }

                    // If any data changed, mark the update
                    if (dataChanged) {
                        updated = true;
                    }
                }
            }

            // If any update was made, write the updated hard file to the new file path
            if (updated) {
                try (FileOutputStream outputStream = new FileOutputStream(newFilePath.toFile())) {
                    hardWorkbook.write(outputStream);
                }
            }
        }
    }

    // Helper method to safely retrieve a string value from a cell
    private String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}
