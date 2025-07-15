package com.example.readword;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.w3c.dom.Node;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;

public class WordToHtmlConverter {

    public static void main(String[] args) throws Exception {
        convertWordToHtml("./22_2023_QH15_518805.docx", "output.html");
    }

    public static void convertWordToHtml(String inputPath, String outputPath) throws Exception {
        try (XWPFDocument doc = new XWPFDocument(new FileInputStream(inputPath));
             FileWriter writer = new FileWriter(outputPath)) {

            writer.write("""
                    <!DOCTYPE html>
                    <html>
                    <head>
                        <meta charset="UTF-8">
                        <title>Document with Track Changes</title>
                        <style>
                            .insertion { color: blue; background-color: #e6f3ff; }
                            .deletion { color: red; text-decoration: line-through; }
                            .comment { background-color: #fffacd; }
                        </style>
                    </head>
                    <body>
                    """);

            processDocument(doc, writer);

            writer.write("</body></html>");
        }
    }

    private static void processDocument(XWPFDocument doc, FileWriter writer) throws IOException {
        // Process paragraphs
        for (XWPFParagraph p : doc.getParagraphs()) {
            writer.write("<p>");
            processParagraph(p, writer);
            writer.write("</p>");
        }

        // Process tables
        for (XWPFTable table : doc.getTables()) {
            writer.write("<table border='1'>");
            for (XWPFTableRow row : table.getRows()) {
                writer.write("<tr>");
                for (XWPFTableCell cell : row.getTableCells()) {
                    writer.write("<td>");
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        processParagraph(p, writer);
                    }
                    writer.write("</td>");
                }
                writer.write("</tr>");
            }
            writer.write("</table>");
        }
    }

    private static void processParagraph(XWPFParagraph p, FileWriter writer) throws IOException {
        for (XWPFRun run : p.getRuns()) {
            CTR ctRun = run.getCTR();
            System.out.println(ctRun.toString());
            String text;

            // Cách 1: Kiểm tra qua XML node
            boolean isDeleted = false;
            boolean isInserted = false;

            Node node = ctRun.getDomNode().getLastChild();

            try {
                text = node.getTextContent();
            } catch (Exception e) {
                text = run.getText(0);
            }

            if (text == null || text.isEmpty()) break;

            if (node.getNodeName().contains("del")) {
                isDeleted = true;
            }
            if (node.getNodeName().contains("ins")) {
                isInserted = true;
            }

            // Cách 2: Kiểm tra thông qua các phương thức khác
            if (!isDeleted && ctRun.getDelTextArray().length > 0) {
                isDeleted = true;
            }
            if (!isInserted && ctRun.getInstrTextArray().length > 0) {
                isInserted = true;
            }

            // Xử lý hiển thị
            if (isDeleted) {
                writer.write("<span class='deletion'>" + escapeHtml(text) + "</span>");
            } else if (isInserted) {
                writer.write("<span class='insertion'>" + escapeHtml(text) + "</span>");
            } else {
                writer.write(escapeHtml(text));
            }
        }
    }

    private static String escapeHtml(String text) {
        if (text == null) return "";
        return text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;");
    }
}