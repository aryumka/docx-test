package aryumka.docxexample;

import java.util.List;
import java.util.Map;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class WordReplacer {

  private XWPFDocument docx;

  public WordReplacer(XWPFDocument docx) {
    this.docx = docx;
  }

  public XWPFDocument replaceTexts(Map<String, String> replacements) {
    replacements.forEach((key, value) -> {
      docx = replaceText(docx, key, value);
    });

    return docx;
  }

  private XWPFDocument replaceText(XWPFDocument doc, String originalText, String updatedText) {
    replaceTextInParagraphs(doc.getParagraphs(), originalText, updatedText);
    for (XWPFTable tbl : doc.getTables()) {
      for (XWPFTableRow row : tbl.getRows()) {
        for (XWPFTableCell cell : row.getTableCells()) {
          replaceTextInParagraphs(cell.getParagraphs(), originalText, updatedText);
        }
      }
    }
    return doc;
  }

  private void replaceTextInParagraphs(List<XWPFParagraph> paragraphs, String originalText, String updatedText) {
    paragraphs.forEach(paragraph -> replaceTextInParagraph(paragraph, originalText, updatedText));
  }

  private void replaceTextInParagraph(XWPFParagraph paragraph, String originalText, String updatedText) {
    List<XWPFRun> runs = paragraph.getRuns();
    for (XWPFRun run : runs) {
      String text = run.getText(0);
      if (text != null && text.contains(originalText)) {
        String updatedRunText = text.replace(originalText, updatedText);
        run.setText(updatedRunText, 0);
      }
    }
  }

}
