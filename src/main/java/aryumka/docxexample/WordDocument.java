package aryumka.docxexample;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Map;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

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

  public void replaceTexts(Map<String, String> replacements) {
    this.docx = this.replacer.replaceTexts(replacements);
  }

  public void concatDocument(String filePath) {
    try {
      XWPFDocument newDocx = new XWPFDocument(new FileInputStream(filePath));
      CTBody newDocxBody = newDocx.getDocument().getBody();
      this.docx.getDocument().addNewBody().set(newDocxBody);
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
}
