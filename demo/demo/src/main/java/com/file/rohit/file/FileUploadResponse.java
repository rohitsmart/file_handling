package com.file.rohit.file;
public class FileUploadResponse {

    private String downloadUrl;

    public FileUploadResponse(String downloadUrl) {
        this.downloadUrl = downloadUrl;
    }

    public String getDownloadUrl() {
        return downloadUrl;
    }

    public void setDownloadUrl(String downloadUrl) {
        this.downloadUrl = downloadUrl;
    }
}

