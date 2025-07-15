package com.example.readword;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.ArrayList;
import java.util.List;

public class LawParser {

    public static class Law {
        public String name;
        public String note;
        public List<Article> articleNumbers = new ArrayList<>();

        @Override
        public String toString() {
            return "Law{name='" + name + "', note='" + note + "', articles=" + articleNumbers + "}";
        }
    }

    public static class Article {
        public String name;
        public String note;
        public String content;
        public List<Clause> clauseNumbers = new ArrayList<>();

        @Override
        public String toString() {
            return "\n  Article{name='" + name + "', note='" + note + "', content='" + content + "', clauses=" + clauseNumbers + "}";
        }
    }

    public static class Clause {
        public String name;
        public String note;
        public String content;
        public List<Section> sectionNumbers = new ArrayList<>();

        @Override
        public String toString() {
            return "\n\n    Clause{name='" + name + "', note='" + note + "', content='" + content + "', sections=" + sectionNumbers + "}";
        }
    }

    public static class Section {
        public String name;
        public String content;

        @Override
        public String toString() {
            return "\n\n\n      Section{name='" + name + "', content='" + content + "'}";
        }
    }

    public static void main(String[] args) throws Exception {
        XWPFDocument docx = new XWPFDocument(OPCPackage.open("./22_2023_QH15_518805.docx"));
        List<IBodyElement> elements = docx.getBodyElements();
        Law law = new Law();
        parseElements(elements, 0, 0, law, null, null, null);
        docx.close();
        System.out.println(law);
    }

    private static int parseElements(List<IBodyElement> elements, int index, int level, Law law,
                                     Article currentArticle, Clause currentClause, Section currentSection) {
        while (index < elements.size()) {
            IBodyElement elem = elements.get(index);

            if (elem instanceof XWPFParagraph para) {
                String style = para.getStyle();
                String text = para.getText().trim();
                int headingLevel = getHeadingLevel(style);

                if (headingLevel > 0) {
                    if (headingLevel <= level) return index;
                    switch (headingLevel) {
                        case 1 -> {
                            law.name = text;
                            return parseElements(elements, index + 1, 1, law, null, null, null);
                        }
                        case 2 -> {
                            Article article = new Article();
                            article.name = text;
                            law.articleNumbers.add(article);
                            index = parseElements(elements, index + 1, 2, law, article, null, null);
                            continue;
                        }
                        case 3 -> {
                            Clause clause = new Clause();
                            clause.name = text.replaceFirst("^Điều\\s+\\d+\\.\\s*", "");
                            if (currentArticle != null)
                                currentArticle.clauseNumbers.add(clause);
                            index = parseElements(elements, index + 1, 3, law, currentArticle, clause, null);
                            continue;
                        }
                        case 4 -> {
                            Section section = new Section();
                            section.name = text.replaceFirst("^(\\d+\\.)\\s*", "");
                            if (currentClause != null)
                                currentClause.sectionNumbers.add(section);
                            index = parseElements(elements, index + 1, 4, law, currentArticle, currentClause, section);
                            continue;
                        }
                    }
                } else {
                    // đoạn văn không có heading -> xử lý nội dung thường
                    String nextStyle = getNextParagraphStyle(elements, index + 1);
                    int nextLevel = getHeadingLevel(nextStyle);

                    boolean isNote = nextLevel > 0 && nextLevel < level;

                    if (currentSection != null) {
                        currentSection.content = concat(currentSection.content, text);
                    } else if (currentClause != null) {
                        if (isNote && currentClause.note == null) currentClause.note = text;
                        else currentClause.content = concat(currentClause.content, text);
                    } else if (currentArticle != null) {
                        if (isNote && currentArticle.note == null) currentArticle.note = text;
                        else currentArticle.content = concat(currentArticle.content, text);
                    } else if (law != null) {
                        law.note = concat(law.note, text);
                    }
                }
            } else if (elem instanceof XWPFTable table) {
                StringBuilder tableText = new StringBuilder("[BẢNG]\n");
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        tableText.append(cell.getText()).append(" | ");
                    }
                    tableText.append("\n");
                }
                String content = tableText.toString();
                if (currentSection != null) currentSection.content = concat(currentSection.content, content);
                else if (currentClause != null) currentClause.content = concat(currentClause.content, content);
                else if (currentArticle != null) currentArticle.content = concat(currentArticle.content, content);
            }
            index++;
        }
        return index;
    }

    private static String getNextParagraphStyle(List<IBodyElement> elements, int fromIndex) {
        for (int i = fromIndex; i < elements.size(); i++) {
            if (elements.get(i) instanceof XWPFParagraph p) {
                return p.getStyle();
            }
        }
        return null;
    }

    private static int getHeadingLevel(String style) {
        if (style != null && style.startsWith("Heading")) {
            try {
                return Integer.parseInt(style.substring(7));
            } catch (NumberFormatException ignored) {
            }
        }
        return 0;
    }

    private static String concat(String oldText, String newText) {
        return oldText == null ? newText : oldText + "\n" + newText;
    }
}
