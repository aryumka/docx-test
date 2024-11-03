package aryumka.docxexample;

import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.springframework.core.io.ClassPathResource;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class TestController {

    @GetMapping
    public String test() {
        try {
            ClassPathResource resource = new ClassPathResource("template.docx");
            var inputPath = Paths.get(resource.getURI()).toString();
            var outputPath = "output.docx";

            WordDocument document = new WordDocument(inputPath);

            var replacements = new HashMap<>(Map.ofEntries(
                    Map.entry("{{A1}}", "2024"),
                    Map.entry("{{A2}}", "유아름"),
                    Map.entry("{{A3}}", "123456-1234567"),
                    Map.entry("{{A4}}", "2024"),
                    Map.entry("{{A5}}", "2023"),
                    Map.entry("{{A6}}", "1,000,000.123"),
                    Map.entry("{{A7}}", "1,000,000.123"),
                    Map.entry("{{A8}}", "1,000,000.123"),
                    Map.entry("{{A9}}", "1,000,000.123"),
                    Map.entry("{{A10}}", "1,000,000.123"),
                    Map.entry("{{A11}}", "1,000,000.123")
            ));

            document.replaceTexts(replacements);
            replacements.clear();

            var rows = document.copyRows(0, 11, 2);
            for (int i = 1; i <= 6; i++) {
                document.replaceTexts(
                        Map.ofEntries(
                                Map.entry("{{C1}}", i + ""),
                                Map.entry("{{C2}}", "서울시 행복구 행복동 123-456"),
                                Map.entry("{{C3}}", "20240101"),
                                Map.entry("{{C4}}", "20240101"),
                                Map.entry("{{C5}}", "10,000"),
                                Map.entry("{{C6}}", "10,000"),
                                Map.entry("{{C7}}", "10%"),
                                Map.entry("{{C8}}", "10%"),
                                Map.entry("{{C9}}", "10,000"),
                                Map.entry("{{C10}}", "10,000"),
                                Map.entry("{{C11}}", "10,000"),
                                Map.entry("{{C12}}", "10,000"),
                                Map.entry("{{C13}}", "100,000,000"),
                                Map.entry("{{C14}}", "100,000,000"),
                                Map.entry("{{C15}}", "100,000,000"),
                                Map.entry("{{C16}}", "100%"),
                                Map.entry("{{C17}}", "100,000,000"),
                                Map.entry("{{C18}}", "100,000,000"),
                                Map.entry("{{C19}}", "100,000,000"),
                                Map.entry("{{C20}}", "100,000,000"),
                                Map.entry("{{C21}}", "100,000,000"),
                                Map.entry("{{C22}}", "100,000,000"),
                                Map.entry("{{C23}}", "100,000,000")
                        )
                );

                if (i != 6) {
                    document.insertRows(0, 11, rows, 13 + (i - 1) * 2);
                }
            }

            document.concatDocument(resource.getInputStream());

            for (var table : document.getDocx().getTables()) {
                table.setTopBorder(XWPFTable.XWPFBorderType.THICK_THIN_LARGE_GAP, 1, 0, "FFFFFF");
                table.setBottomBorder(XWPFTable.XWPFBorderType.THICK_THIN_LARGE_GAP, 1, 0, "FFFFFF");
                table.setLeftBorder(XWPFTable.XWPFBorderType.THICK_THIN_LARGE_GAP, 1, 0, "FFFFFF");
                table.setRightBorder(XWPFTable.XWPFBorderType.THICK_THIN_LARGE_GAP, 1, 0, "FFFFFF");
            }

            document.save(outputPath);
            document.convertPdf("output.pdf");
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        return "Hello, World!";
    }

}


