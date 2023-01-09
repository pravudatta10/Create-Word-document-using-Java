
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import java.math.BigInteger;
 

 
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
 
 
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

// For Header And Footer 
class HeaderFooter
 	{	 
		public   void headerFooter(XWPFDocument  document) 
		{			
		  // create header-footer
		  XWPFHeaderFooterPolicy headerFooterPolicy = document.getHeaderFooterPolicy();
		  if (headerFooterPolicy == null) headerFooterPolicy = document.createHeaderFooterPolicy();

		  // create header start
		  XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);

		  XWPFParagraph paragraph = header.createParagraph();
		  paragraph.setAlignment(ParagraphAlignment.CENTER);
		  XWPFRun run = paragraph.createRun();  
		  run.setBold(true);
		  run.setColor("000000");
		  run.setFontSize(18);
		  run.setText("Vehicle Details");
		  

		  // create footer start
		  
		  XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
		  XWPFParagraph paragraph1 = footer.createParagraph();
		  paragraph1.setAlignment(ParagraphAlignment.CENTER);
		  XWPFRun run1 = paragraph1.createRun();  
	
		  
		  run1.setText("");
		  
		  XWPFTable table = footer.createTable(2, 3);
		  table.setWidth((int) (6.5 * 1540));
		 
		  table.setTableAlignment(TableRowAlign.CENTER);
		  table.setTopBorder(XWPFTable.XWPFBorderType.SINGLE, 5,0, "FF0000");
		  table.setBottomBorder(XWPFTable.XWPFBorderType.SINGLE, 5, 0, "FF0000");
		  table.setLeftBorder(XWPFTable.XWPFBorderType.SINGLE, 5, 0, "FF0000");
		  table.setRightBorder(XWPFTable.XWPFBorderType.SINGLE, 5, 0, "FF0000");
		  table.setInsideHBorder(XWPFTable.XWPFBorderType.SINGLE, 5, 0, "FF0000");
		  table.setInsideVBorder(XWPFTable.XWPFBorderType.SINGLE, 5, 0, "FF0000");
			
		
			// write to first row, first column
			XWPFParagraph p1 = table.getRow(0).getCell(0).getParagraphs().get(0);		
			XWPFRun r1 = p1.createRun();			 	 
			r1.setText("Manufacturer");
			r1.setColor("FF0000");
			r1.setFontSize(12);
			
			// write to first row, second column
			XWPFParagraph p2 = table.getRow(0).getCell(1).getParagraphs().get(0);			 
			XWPFRun r2 = p2.createRun();			 
			r2.setText("Sheet No :  							");
			r2.setColor("FF0000");
			r2.setFontSize(12);
			
			// write to first row, third column
			XWPFParagraph p3 = table.getRow(0).getCell(2).getParagraphs().get(0);		
			XWPFRun r3 = p3.createRun();		
			r3.setText("Test Agency :");
			r3.setColor("FF0000");
			r3.setFontSize(12);				
			table.getRow(1).getCell(0).setText("");
			 
		 
			XWPFParagraph p5 = table.getRow(1).getCell(0).getParagraphs().get(0);		
			XWPFRun r5 = p5.createRun();		
			r5.setText("Name: 																Designation:"
					+ "");
			r5.setColor("FF0000");
			r5.setFontSize(12);
			
			XWPFParagraph p6 = table.getRow(1).getCell(1).getParagraphs().get(0);			
			XWPFRun r6 = p6.createRun();		
			r6.setText("Date :");
			r6.setColor("FF0000");
			r6.setFontSize(12);
			
			XWPFParagraph p7 = table.getRow(1).getCell(2).getParagraphs().get(0);			 
			XWPFRun r7 = p7.createRun();		 
			r7.setText("Name: 		Designation:");
			r7.setColor("FF0000");
			r7.setFontSize(12); 
			  
			
			//Add Page Number
				 XWPFParagraph pageno = footer.createParagraph();
				 pageno.setAlignment(ParagraphAlignment.RIGHT);
				  XWPFRun page = pageno.createRun();
				  XWPFRun pageNo = pageno.createRun();    
				  page.setText("Page | ");
				  page.setColor("000000");
				  page.setFontSize(12);
				  pageNo.getCTR().addNewPgNum();				 
				  pageNo.setColor("FF0000");
				  pageNo.setFontSize(12);
			  
			
			  
		  CTSectPr sectPr = document.getDocument().getBody().getSectPr();
		  if (sectPr == null) sectPr = document.getDocument().getBody().addNewSectPr();
		  CTPageMar pageMar = sectPr.getPgMar();
		  if (pageMar == null) pageMar = sectPr.addNewPgMar();
		  pageMar.setLeft(BigInteger.valueOf(620)); //720 TWentieths of an Inch Point (Twips) = 720/20 = 36 pt = 36/72 = 0.5"
		  pageMar.setRight(BigInteger.valueOf(720));
		  pageMar.setTop(BigInteger.valueOf(1000)); //1440 Twips = 1440/20 = 72 pt = 72/72 = 1"
		  pageMar.setBottom(BigInteger.valueOf(1000));
		  pageMar.setHeader(BigInteger.valueOf(400)); //45.4 pt * 20 = 908 = 45.4 pt header from top
		  pageMar.setFooter(BigInteger.valueOf(300)); //28.4 pt * 20 = 568 = 28.4 pt footer from bottom
		 }
	}



public class WordDocument  
{
	static XWPFTable table;

	public static void BodyTable(XWPFDocument document) {

		String[] column1 = { "1.0", "1.1.", "1.1.1.", "1.1.2."  };
		int len = column1.length;
		table = document.createTable(len, 3);
		table.setWidth((int) (6.5 * 1540));
		table.setTableAlignment(TableRowAlign.CENTER);		 
	 
	
		for (int i = 1; i <len; i++) 
		{
			// write to first row, first column
			 XWPFTableCell cell = table.getRow(i).getCell(0);
			 cell.setVerticalAlignment(XWPFVertAlign.CENTER );
			 cell.setText(column1[i]);			 
			 cell.getCTTc().addNewTcPr().addNewShd().setFill("FFFFCC");
			 
		}
		String[] column2 = 
			{
				"Description of energy converters and power plant. (In the case of a vehicle that can run either on petrol, diesel, etc., or also in combination with another fuel, items shall be repeated.",
				"Manufacturer's engine model code (as marked on the engine, or other means of identification):",
				"Location of engine model code and engine serial number on block / crankcase (drawing to be attached)",
				"Internal combustion engine:" };
		
		for (int i = 1; i < column2.length; i++) {
			// write to first row, second column
			XWPFTableCell cell = table.getRow(i).getCell(1);
			 cell.setVerticalAlignment(XWPFVertAlign.CENTER );
			 cell.setText(column2[i]);
			 
		}

		String[] column3 = { " ", " ", " ", " "  };
		for (int i = 1; i < column3.length; i++) 
		{
			// write to first row, third column
			XWPFTableCell cell = table.getRow(i).getCell(2);
			 cell.setVerticalAlignment(XWPFVertAlign.CENTER );
			 cell.setText(column3[i]);
		
		}
		//For Create a New Page  
		
		XWPFParagraph paragraph = document.createParagraph();
		XWPFRun run = paragraph.createRun();		 
		run.addBreak(BreakType.PAGE);
		
		//ADD IMAGE TO the DOC
			XWPFParagraph p3 = table.getRow(2).getCell(2).getParagraphs().get(0);
			 XWPFRun r4 = p3.createRun();
			 
			p3.setVerticalAlignment(TextAlignment .CENTER);
			File image = new File("Assets\\Elephant.png");
			FileInputStream imageData = null ;
			try 
			{
				imageData = new FileInputStream(image);
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			int imageType = XWPFDocument.PICTURE_TYPE_PNG;
			String imageFileName = image.getName();
			int width = 70;
			int height = 30;
			try 
			{
				r4.addPicture(imageData, imageType, imageFileName, Units.toEMU(width), Units.toEMU(height));
			} catch (InvalidFormatException | IOException e) 
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		
	 //Merge the Top Cell and make it heading 
	
	static void mergeCellHorizontally(XWPFTable table, int row, int fromCol, int toCol) 
	{
		XWPFTableCell cell = table.getRow(row).getCell(fromCol);
		// Try getting the TcPr. Not simply setting an new one every time.
		CTTcPr tcPr = cell.getCTTc().getTcPr();
		if (tcPr == null)
			tcPr = cell.getCTTc().addNewTcPr();
		// The first merged cell has grid span property set
		if (tcPr.isSetGridSpan()) {
			tcPr.getGridSpan().setVal(BigInteger.valueOf(toCol - fromCol + 1));
		} else {
			tcPr.addNewGridSpan().setVal(BigInteger.valueOf(toCol - fromCol + 1));
		}
		// Cells which join (merge) the first one, must be removed
		for (int colIndex = toCol; colIndex > fromCol; colIndex--) {
			table.getRow(row).removeCell(colIndex); // use only this for apache poi versions greater than 3
			// table.getRow(row).getCtRow().removeTc(colIndex); // use this for apache poi
			// versions up to 3
			// table.getRow(row).removeCell(colIndex);
		}
		 XWPFParagraph header = table.getRow(0).getCell(0).getParagraphs().get(0);
		  header.setAlignment(ParagraphAlignment.CENTER); 
		  XWPFRun runhead = header.createRun();
		  runhead.setText(
		  "TECHNICAL SPECIFICATIONS FOUR WHEELER ");
		  runhead.setFontFamily("Caiber");
		  runhead.setBold(true);
		  runhead.setFontSize(14);
	}

	public static void main(String[] args) throws Exception 
	{
		 
		// Blank Document
		XWPFDocument document = new XWPFDocument();
		// Write the Document in file system
		
		 
		  HeaderFooter hf = new HeaderFooter(); hf.headerFooter(document);
		 
		BodyTable(document);
		mergeCellHorizontally(table, 0, 0, 2);

		// Add Image Here
		XWPFParagraph paragraph = document.createParagraph();
		XWPFRun run = paragraph.createRun();
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		// specifying images path
		File image = new File("Assets\\Elephant.png");
		FileInputStream imageData = new FileInputStream(image);
		// Step 5: Retrieving the image file name and image type				
		int imageType = XWPFDocument.PICTURE_TYPE_PNG;
		String imageFileName = image.getName();
		// Step 6: Setting the width and height of the image in pixels.
		int width = 250;
		int height = 200;
		// Step 7: Adding the picture using the addPicture()
		run.addPicture(imageData, imageType, imageFileName, Units.toEMU(width), Units.toEMU(height));
	
		// Image Part End
		
		String outputdoc=" OUTPUT\\WordReport.docx";
		FileOutputStream out = new FileOutputStream(new File(outputdoc));		
	 
        document.write(out);
        out.close();
       System.out.println("Done");
	}

}
