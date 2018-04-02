package com.dabanin.reporter.controller;

import com.dabanin.reporter.service.ReportService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;


@RestController
@RequestMapping("/report")
public class ReportController {

    private ReportService reportService;

    public ReportController(ReportService reportService) {
        this.reportService = reportService;
    }

    @GetMapping("/{id}")
    public ResponseEntity<byte[]> getReportById(@PathVariable Long id) throws IOException {
        HttpHeaders headers = buildHeaders(id, "reportWithId_");
        return new ResponseEntity<>(reportService.getReport(id), headers, HttpStatus.OK);
    }

    @GetMapping("/table/{id}")
    public ResponseEntity<byte[]> getTableReportById(@PathVariable Long id) throws IOException {
        HttpHeaders headers = buildHeaders(id, "tableReportWithId_");
        return new ResponseEntity<>(reportService.getTableReportService(id), headers, HttpStatus.OK);
    }

    private HttpHeaders buildHeaders(Long id, String reportName) {
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        String filename = reportName + id + ".docx";
        headers.setContentDispositionFormData(filename, filename);
        return headers;
    }


}
