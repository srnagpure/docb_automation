import java.io.FileInputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSignedTwipsMeasure;


public class WordDocumentParser {
    public static void main(String[] args) throws Exception {
        // Load the Word document
        String filePath = "generated_doc.docx"; // Replace with your document path
        XWPFDocument document = new XWPFDocument(new FileInputStream(filePath));

        System.out.println("Document Parsing Started...\n");
        
        
        // Get section properties
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        
        // Retrieve margin values (in twips, 1 inch = 1440 twips)
        double topMargin = (sectPr.getPgMar().getTop().doubleValue() / 1440) * 2.54 ;       // Convert to inches and then cms
        double bottomMargin = (sectPr.getPgMar().getBottom().doubleValue() / 1440) * 2.54;
        double leftMargin = (sectPr.getPgMar().getLeft().doubleValue() / 1440) * 2.54;
        double rightMargin = (sectPr.getPgMar().getRight().doubleValue()/ 1440) * 2.54;
        
        System.out.println("Top Margin ::"+ topMargin +" cm");
        System.out.println("Bottom Margin ::"+ bottomMargin +" cm");
        System.out.println("Left Margin ::"+ leftMargin +" cm");
        System.out.println("Right Margin ::"+ rightMargin +" cm");

        // Iterate through all paragraphs
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            System.out.println("Paragraph Text: " + paragraph.getText());

            // Print line spacing (leading space)
            double lineSpacing = paragraph.getSpacingAfter();
            System.out.println("  Line Spacing: " + (lineSpacing > 0 ? lineSpacing : "Default"));

            // Iterate through all runs in the paragraph
            for (XWPFRun run : paragraph.getRuns()) {
                System.out.println("    Run Text: " + run.getText(0));
                
                // Print font details
                String fontName = run.getFontFamily();
                int fontSize = run.getFontSize();
                System.out.println("      Font Name: " + (fontName != null ? fontName : "Default"));
                System.out.println("      Font Size: " + (fontSize > 0 ? fontSize : "Default"));

                // Check bold and italic properties
                System.out.println("      Is Bold: " + run.isBold());
                System.out.println("      Is Italic: " + run.isItalic());

                // Print character spacing (tracking space)
                if (run.getCTR() != null && run.getCTR().getRPr() != null && run.getCTR().getRPr().getSpacing() != null) {
                    CTSignedTwipsMeasure spacing = run.getCTR().getRPr().getSpacing();
                    int characterSpacing = spacing.getVal().intValue(); // In twips
                    System.out.println("      Character Spacing: " + characterSpacing + " twips");
                } else {
                    System.out.println("      Character Spacing: Default");
                }
            }

            System.out.println(); // Blank line after each paragraph
        }

        
        // Iterate through all tables
        List<XWPFTable> tables = document.getTables();
        for (XWPFTable table : tables) {
        	System.out.println("Table Width: " + table.getWidth());
        	
            for (XWPFTableRow row : table.getRows()) {
            	System.out.println("Row Height: " + row.getHeight());
                for (XWPFTableCell cell : row.getTableCells()) {
                    System.out.println("Cell Text: " + cell.getText());
                }
            }
        }

        // Iterate through all footers
        List<XWPFFooter> footers = document.getFooterList();
        for (XWPFFooter footer : footers) {
            List<XWPFParagraph> footerParagraphs = footer.getParagraphs();
            for (XWPFParagraph footerParagraph : footerParagraphs) {
                System.out.println("Footer Paragraph Text: " + footerParagraph.getText());
            }
        }

        // Iterate through all headers
        List<XWPFHeader> headers = document.getHeaderList();
        for (XWPFHeader header : headers) {
            List<XWPFParagraph> headerParagraphs = header.getParagraphs();
            for (XWPFParagraph headerParagraph : headerParagraphs) {
                System.out.println("Header Paragraph Text: " + headerParagraph.getText());
            }
        }
        
        System.out.println("Document Parsing Completed.");
    }
}
