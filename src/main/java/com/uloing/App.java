package com.uloing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.StringWriter;

import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.w3c.dom.Node;

/**
 * Hello world!
 */
public final class App {
  private App() {
  }

  /**
   * Says hello to the world.
   * 
   * @param args The arguments of the program.
   */
  public static void main(String[] args) {
    try (FileInputStream fis = new FileInputStream(new File("./document.docx"));
        FileOutputStream fos = new FileOutputStream(new File("./result.docx"));) {
      XWPFDocument doc = new XWPFDocument(fis);

      XWPFTable table = doc.createTable();
      XWPFTableRow row;
      XWPFTableCell cell;
      XWPFParagraph paragraph;
      XWPFRun run;

      // align table
      CTTblPr tblPr = table.getCTTbl().getTblPr();
      CTJc jc = tblPr.getJc();
      if (jc == null) {
        jc = tblPr.addNewJc();
      }
      jc.setVal(STJc.CENTER);
      tblPr.setJc(jc);

      // Header
      String[] tableHeaders = { "Vertragskonto", "Abnahmestelle", "Marktlokation", "Zählverfahren", "Menge [kWh]" };

      row = table.getRow(0);
      for (int i = 0; i < tableHeaders.length; i++) {
        if (i == 0) {
          cell = row.getCell(0);
        } else {
          cell = row.addNewTableCell();
        }

        paragraph = cell.getParagraphs().get(0); // cell.addParagraph();

        run = paragraph.createRun();
        run.setText(tableHeaders[i]);
        run.setBold(true);
        run.setUnderline(UnderlinePatterns.SINGLE);
        run.setFontSize(10);
        run.setFontFamily("Arial");
      }

      System.out.println(nodeToString(table.getCTTbl().getDomNode()));

      doc.write(fos);
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static String nodeToString(Node node) {
    StringWriter sw = new StringWriter();
    try {
      Transformer t = TransformerFactory.newInstance().newTransformer();
      t.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");
      t.setOutputProperty(OutputKeys.INDENT, "yes");
      t.transform(new DOMSource(node), new StreamResult(sw));
    } catch (TransformerException te) {
      System.out.println("nodeToString Transformer Exception");
    }
    return sw.toString();
  }

}
