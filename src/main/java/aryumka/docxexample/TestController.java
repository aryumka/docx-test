package aryumka.docxexample;

import java.util.Map;
import org.springframework.core.io.ClassPathResource;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class TestController {

  @GetMapping
  public String test() {
    ClassPathResource resource = new ClassPathResource("input.docx");
    var outputPath = "output.docx";

    WordDocument document = new WordDocument(resource.getPath());
    document.replaceTexts(
        Map.ofEntries(
            Map.entry("{{C1}}", "2024"),
            Map.entry("{{C2}}", "유아름"),
            Map.entry("{{C3}}", "123456-1234567"),
            Map.entry("{{C4}}", "1,000,000.123"),
            Map.entry("{{C5}}", "1,000,000.123"),
            Map.entry("{{C6}}", "1,000,000.123"),
            Map.entry("{{C7}}", "1,000,000.123"),
            Map.entry("{{C8}}", "1,000,000.123"),
            Map.entry("{{C9}}", "1,000,000.123"),
            Map.entry("{{C10}}", "1"),
            Map.entry("{{C11}}", "서울특별시 용산구 효창원로66길 13, 201호")
        )
    );

    document.concatDocument(resource.getPath());

    document.save(outputPath);

    return "Hello, World!";
  }

}


