package com.dabanin.reporter.service;

import com.dabanin.reporter.entity.PersonInfo;
import com.dabanin.reporter.entity.TableInfo;
import com.dabanin.reporter.repository.PersonInfoRepository;
import com.dabanin.reporter.repository.TableInfoRepository;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Map;

@Service
public class ReportService {

    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd.MM.yyyy");

    private PersonInfoRepository personInfoRepository;
    private TableInfoRepository tableInfoRepository;
    private WordProcessService wordProcessService;

    public ReportService(PersonInfoRepository personInfoRepository, TableInfoRepository tableInfoRepository, WordProcessService wordProcessService) {
        this.personInfoRepository = personInfoRepository;
        this.tableInfoRepository = tableInfoRepository;
        this.wordProcessService = wordProcessService;
    }

    public byte[] getReport(Long id) throws IOException {
        PersonInfo personInfo = personInfoRepository.findById(id).orElseThrow(RuntimeException::new);
        Map<String, String> variablesMap = new HashMap<>();
        variablesMap.put("$CompanyName", personInfo.getCompanyName());
        variablesMap.put("$DirectorName", personInfo.getDirectorName());
        variablesMap.put("$Position", personInfo.getPosition());
        variablesMap.put("$Name", personInfo.getName());
        variablesMap.put("$DateFrom", personInfo.getDateFrom().format(DATE_FORMATTER));
        variablesMap.put("$CurrentDate", LocalDate.now().format(DATE_FORMATTER));
        variablesMap.put("$DayCount", personInfo.getDaysCount().toString());
        XWPFDocument doc = wordProcessService.processDoc(variablesMap);
        try (ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()) {
            doc.write(byteArrayOutputStream);
            return byteArrayOutputStream.toByteArray();
        }
    }

    public byte[] getTableReportService(Long id) throws IOException {
        TableInfo tableInfo = tableInfoRepository.findById(id).orElseThrow(RuntimeException::new);
        Map<String, String> valueMap = new HashMap<>();
        valueMap.put("ИНН", tableInfo.getInn());
        valueMap.put("Полное наименвоание", tableInfo.getFullName());
        valueMap.put("Краткое наименвоание", tableInfo.getShortName());
        valueMap.put("ОПФ", tableInfo.getTypeName());

        Map<String, String> variablesMap = new HashMap<>();
        variablesMap.put("$CurrentDate", LocalDate.now().format(DATE_FORMATTER));
        variablesMap.put("$DocNumber", tableInfo.getId().toString());
        variablesMap.put("$CompanyName", tableInfo.getFullName());
        XWPFDocument tableDoc = wordProcessService.createTableDoc(valueMap, variablesMap);
        try (ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()) {
            tableDoc.write(byteArrayOutputStream);
            return byteArrayOutputStream.toByteArray();
        }
    }

    public byte[] getReportAsPdf(Long id) throws IOException {
        return wordProcessService.convertToPDF(getReport(id));
    }
}
