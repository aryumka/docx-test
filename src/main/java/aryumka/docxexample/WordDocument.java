package aryumka.docxexample;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.docx4j.Docx4J;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.springframework.core.io.ClassPathResource;

public class WordDocument {
    private XWPFDocument docx;
    private final WordReplacer replacer;

    public WordDocument(String filePath) {
        try {
            this.docx = new XWPFDocument(new FileInputStream(filePath));
            this.replacer = new WordReplacer(docx);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public WordDocument(InputStream stream) {
        try {
            this.docx = new XWPFDocument(stream);
            this.replacer = new WordReplacer(docx);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public XWPFDocument getDocx() {
        return this.docx;
    }

    public void replaceTexts(Map<String, String> replacements) {
        this.docx = this.replacer.replaceTexts(replacements);
        this.commitTables();
    }

    public List<XWPFTableRow> copyRows(int tableIndex, int rowIndex, int count) {
        XWPFTable table = this.docx.getTables().get(tableIndex);

        var copiedRows = new ArrayList<XWPFTableRow>();
        for (int i = 0; i < count; i++) {
            var row = table.getRow(rowIndex + i);
            copiedRows.add(new XWPFTableRow((CTRow) row.getCtRow().copy(), table));
        }

        return copiedRows;
    }

    public void insertRows(int tableIndex, int rowIndex, List<XWPFTableRow> rows, int insertIndex) {
        XWPFTable table = this.docx.getTables().get(tableIndex);

        for (int i = 0; i < rows.size(); i++) {
            var row = new XWPFTableRow((CTRow) rows.get(i).getCtRow().copy(), table);
            table.addRow(row, insertIndex + i);
        }

        this.commitTables();
    }

    public void duplicateRows(int tableIndex, int rowIndex, int count, int insertIndex) {
        XWPFTable table = this.docx.getTables().get(tableIndex);

        var copiedRows = new ArrayList<XWPFTableRow>();
        for (int i = 0; i < count; i++) {
            var row = table.getRow(rowIndex + i);
            copiedRows.add(new XWPFTableRow((CTRow) row.getCtRow().copy(), table));
        }

        for (int i = 0; i < count; i++) {
            table.addRow(copiedRows.get(i), insertIndex + i);
        }

        this.commitTables();
    }

    public void commitTables() {
        this.docx.getTables().forEach(table -> {
            int rowNr = 0;
            for (XWPFTableRow tableRow : table.getRows()) {
                table.getCTTbl().setTrArray(rowNr++, tableRow.getCtRow());
            }
        });
    }

    public void concatDocument(String filePath) {
        try {
            XWPFDocument newDocx = new XWPFDocument(new FileInputStream(filePath));
            CTBody newDocxBody = newDocx.getDocument().getBody();
            this.docx.getDocument().addNewBody().set(newDocxBody);
            this.commitTables();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public void concatDocument(InputStream second) {
        try {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            this.docx.write(out);
            InputStream first = new ByteArrayInputStream(out.toByteArray());
            out.close();

            var stream = WordMerger.merge(List.of(first, second));
            this.docx = new XWPFDocument(Objects.requireNonNull(stream));
            this.commitTables();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public void save(String outputPath) {
        try {
            this.docx.write(new FileOutputStream(outputPath));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public void convertPdf(String outputPath) {
        try {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            this.docx.write(out);
            InputStream in = new ByteArrayInputStream(out.toByteArray());
            out.close();

            Mapper fontMapper = new IdentityPlusMapper();
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(in);

            ClassPathResource fontResource = new ClassPathResource("NanumGothic.ttf");
            PhysicalFonts.addPhysicalFonts("Nanum Gothic", fontResource.getURI());
            var font = PhysicalFonts.get("Nanum Gothic");
            fontMapper.put("Nanum Gothic", font);
            wordMLPackage.setFontMapper(fontMapper);

            FileOutputStream os = new FileOutputStream(outputPath);
            Docx4J.toPDF(wordMLPackage, os);
            os.flush();
            os.close();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
