package com.dabanin.reporter.service;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;
import java.util.Map;

@Service
public class WordProcessService {

    XWPFDocument processDoc(Map<String, String> variableMap) throws IOException {
        XWPFDocument document = new XWPFDocument(ClassLoader.getSystemResourceAsStream("Template.docx"));
        replace(document, variableMap);
        return document;
    }

    byte[] convertToPDF(byte[] docBytes) throws IOException {
        XWPFDocument document;
        try (ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(docBytes)) {
            document = new XWPFDocument(byteArrayInputStream);
        }
        PdfOptions options = PdfOptions.create();
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            PdfConverter.getInstance().convert(document, out, options);
            return out.toByteArray();
        }
    }

    XWPFDocument createTableDoc(Map<String, String> valueMap, Map<String, String> variableMap) throws IOException {
        XWPFDocument document = new XWPFDocument(ClassLoader.getSystemResourceAsStream("TableTemplate.docx"));
        replace(document, variableMap);
        XWPFParagraph para = document.createParagraph();
        para.createRun();
        XWPFTable table = document.createTable();
        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText("Поле");
        tableRowOne.addNewTableCell().setText("Значение");
        valueMap.forEach((param, value) -> addNewRow(table, param, value));
        resizeTable(table);
        return document;
    }

    private void resizeTable(XWPFTable table) {
        for(int x = 0;x < table.getNumberOfRows(); x++){
            XWPFTableRow row = table.getRow(x);
            int numberOfCell = row.getTableCells().size();
            for(int y = 0; y < numberOfCell ; y++){
                XWPFTableCell cell = row.getCell(y);
                cell.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4550L));
            }
        }
    }

    private void addNewRow(XWPFTable table, String param, String value) {
        XWPFTableRow tableRowTwo = table.createRow();
        tableRowTwo.getCell(0).setText(param);
        tableRowTwo.getCell(1).setText(value);
    }

    private void replace(XWPFDocument document, Map<String, String> variableMap) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        paragraphs.forEach(xwpfParagraph -> variableMap
                .forEach((variable, value) -> replace(xwpfParagraph, variable, value)));
    }

    private void replace(XWPFParagraph paragraph, String searchValue, String replacement) {
        if (hasReplaceableItem(paragraph.getText(), searchValue)) {
            String replacedText = StringUtils.replace(paragraph.getText(), searchValue, replacement);
            removeAllRuns(paragraph);
            insertReplacementRuns(paragraph, replacedText);
        }
    }

    private void insertReplacementRuns(XWPFParagraph paragraph, String replacedText) {
        String[] replacementTextSplitOnCarriageReturn = StringUtils.split(replacedText, "\n");
        for (int j = 0; j < replacementTextSplitOnCarriageReturn.length; j++) {
            String part = replacementTextSplitOnCarriageReturn[j];
            XWPFRun newRun = paragraph.insertNewRun(j);
            newRun.setText(part);
            if (j + 1 < replacementTextSplitOnCarriageReturn.length) {
                newRun.addCarriageReturn();
            }
        }
    }

    private void removeAllRuns(XWPFParagraph paragraph) {
        int size = paragraph.getRuns().size();
        for (int i = 0; i < size; i++) {
            paragraph.removeRun(0);
        }
    }

    private boolean hasReplaceableItem(String runText, String searchValue) {
        return StringUtils.contains(runText, searchValue);
    }



}
