package com.file.rohit.file;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import lombok.RequiredArgsConstructor;

import org.springframework.core.io.Resource;
import org.springframework.core.io.UrlResource;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.UUID;

@RestController
@RequestMapping("/api/files")
@RequiredArgsConstructor
public class FileController {

    @Value("${file.upload-dir}")
    private String uploadDir;
    private final FileProcessingService fileProcessingService;
    private final FileCompareService fileCompareService;


    @PostMapping("/upload")
    public ResponseEntity<?> uploadCsvAndGetExcel(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty() || !file.getOriginalFilename().endsWith(".csv")) {
            return ResponseEntity.status(HttpStatus.BAD_REQUEST)
                    .body("Invalid file. Please upload a valid CSV file.");
        }
        try {
            String fileName = UUID.randomUUID().toString() + ".xlsx";
            Path filePath = Paths.get(uploadDir, fileName);
            fileProcessingService.processCsvToExcelAndSave(file, filePath);
            String fileDownloadUri = "/api/files/download/" + fileName;
            return ResponseEntity.ok()
                    .body(new FileUploadResponse(fileDownloadUri));
        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("Error processing file: " + e.getMessage());
        }
    }

    @GetMapping("/download/{fileName}")
    public ResponseEntity<Resource> downloadFile(@PathVariable String fileName) {
        try {
            Path filePath = Paths.get(uploadDir, fileName);
            Resource resource = new UrlResource(filePath.toUri());
            if (resource.exists()) {
                HttpHeaders headers = new HttpHeaders();
                headers.add("Content-Disposition", "attachment; filename=" + fileName);
                headers.setContentType(MediaType.parseMediaType(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
                return ResponseEntity.ok()
                        .headers(headers)
                        .body(resource);
            } else {
                return ResponseEntity.status(HttpStatus.NOT_FOUND)
                        .body(null);
            }
        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body(null);
        }
    }

    @PostMapping("/compare-and-update")
    public ResponseEntity<?> compareAndUpdateFiles(
            @RequestParam("softFile") MultipartFile softFile,
            @RequestParam("hardFile") MultipartFile hardFile) {
        try {
            Path newFilePath = Paths.get(uploadDir, UUID.randomUUID().toString() + ".xlsx");
            fileCompareService.compareAndUpdateFiles(softFile, hardFile, newFilePath);
            String fileDownloadUri = "/api/files/download/" + newFilePath.getFileName().toString();
            return ResponseEntity.ok()
                    .body(new FileUploadResponse(fileDownloadUri));      
          } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("Error processing files: " + e.getMessage());
        }
    }
}
