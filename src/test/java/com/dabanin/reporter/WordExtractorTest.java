package com.dabanin.reporter;


import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.JUnit4;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;


@RunWith(JUnit4.class)
public class WordExtractorTest {

    @Test
    public void test() throws IOException{
        XWPFDocument doc = new XWPFDocument(ClassLoader.getSystemResourceAsStream("Template.docx"));
        for (XWPFParagraph p : doc.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    if (text != null && text.contains("$Name")) {
                        text = text.replace("$Name", "Петров Петр Васильевич");
                        r.setText(text, 0);
                    }
                }
            }
        }
//        for (XWPFTable tbl : doc.getTables()) {
//            for (XWPFTableRow row : tbl.getRows()) {
//                for (XWPFTableCell cell : row.getTableCells()) {
//                    for (XWPFParagraph p : cell.getParagraphs()) {
//                        for (XWPFRun r : p.getRuns()) {
//                            String text = r.getText(0);
//                            if (text != null && text.contains("переводов")) {
//                                text = text.replace("переводов", "haystack");
//                                r.setText(text, 0);
//                            }
//                        }
//                    }
//                }
//            }
//        }
        doc.write(new FileOutputStream("output.docx"));
    }
}
