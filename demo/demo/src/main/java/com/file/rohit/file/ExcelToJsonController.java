package com.file.rohit.file;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;

@RestController
@RequestMapping("/api/excel")
public class ExcelToJsonController {

    @PostMapping("/convert")
    public Map<String, Object> convertExcelToJson(@RequestParam("file") MultipartFile file) {
        Map<String, Object> result = new HashMap<>();
        List<Map<String, Object>> dataList = new ArrayList<>();

        try {
            // Load the Excel file
            InputStream inputStream = file.getInputStream();
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheetAt(0);  // Assuming we are working with the first sheet

            // Get headers
            Row headerRow = sheet.getRow(0);  // First row as headers
            if (headerRow == null) {
                result.put("error", "No header row found in the Excel file.");
                return result;
            }

            List<String> headers = new ArrayList<>();
            for (Cell headerCell : headerRow) {
                headers.add(headerCell.getStringCellValue());
            }

            // Iterate through each row and convert it to a map
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    continue;  // Skip empty rows
                }

                Map<String, Object> dataMap = new HashMap<>();
                for (int j = 0; j < headers.size(); j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        switch (cell.getCellType()) {
                            case STRING:
                                dataMap.put(headers.get(j), cell.getStringCellValue());
                                break;
                            case NUMERIC:
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    dataMap.put(headers.get(j), cell.getDateCellValue());
                                } else {
                                    dataMap.put(headers.get(j), cell.getNumericCellValue());
                                }
                                break;
                            case BOOLEAN:
                                dataMap.put(headers.get(j), cell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                dataMap.put(headers.get(j), cell.getCellFormula());
                                break;
                            default:
                                dataMap.put(headers.get(j), null);
                        }
                    }
                }
                dataList.add(dataMap);
            }

            workbook.close();
            result.put("data", dataList);
            return result;
        } catch (Exception e) {
            result.put("error", "Failed to process the Excel file: " + e.getMessage());
            return result;
        }
    }
}
