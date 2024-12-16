import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSignedTwipsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;

import java.io.FileInputStream;
import java.math.BigInteger;

public class WordDocTest {
    public static void main(String[] args) throws Exception {
        FileInputStream fis = new FileInputStream("generated_doc.docx");
        XWPFDocument doc = new XWPFDocument(fis);
        
        
        // Get section properties
        CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
        
        
        
        
        // Retrieve margin values (in twips, 1 inch = 1440 twips)
        double topMargin = sectPr.getPgMar().getTop().doubleValue() / 1440 ;       // Convert to inches
        double bottomMargin = sectPr.getPgMar().getBottom().doubleValue() / 1440 ;
        double leftMargin = sectPr.getPgMar().getLeft().doubleValue() / 1440;
        double rightMargin = sectPr.getPgMar().getRight().doubleValue()/ 1440;
        
        System.out.println("Margins ::"+topMargin +" :: "+bottomMargin +" :: "+leftMargin +" :: ");
        
        

        // Extract content
        StringBuilder content = new StringBuilder();
        for (XWPFParagraph paragraph : doc.getParagraphs()) {
            
        }

        // Validate content
        // assert content.toString().contains("Expected text here") : "Validation failed";
        if(content.toString().contains("Expected text here")) {
        	System.out.println("Pass");
        	
        }else {
        	System.out.println("Fail");
        }
        	
        
        	
        //doc.close();
    }

    private static void getBodyInfo(IBodyElement bodyElement) {
    	if (BodyElementType.PARAGRAPH == bodyElement.getElementType()) {
    		BodyType partType = bodyElement.getPartType();
    		if(BodyType.DOCUMENT == partType) {
    			
    		}
    	}
    	
    }
    
    
    
    private static void getRunInfo(XWPFParagraph paragraph) {
		for (XWPFRun run : paragraph.getRuns()) {
			
			CTSignedTwipsMeasure spacing = run.getCTR().getRPr().getSpacing();
		    // Access character spacing
		    if (spacing != null) {
		        int characterSpacing = spacing.getVal().intValue(); // Value is in twips
		        System.out.println("Character spacing for "+""+ run.getText(0) +"::"+characterSpacing);
		        // assert characterSpacing == expectedSpacing : "Character spacing mismatch";
		    }
		    System.out.println("Font Name ::"+ run.getFontFamily());
		    
		    XWPFParagraph paragraph2 = run.getParagraph();
		    if(paragraph2 != null)
		    	getRunInfo(paragraph2);
		}
	}
}
