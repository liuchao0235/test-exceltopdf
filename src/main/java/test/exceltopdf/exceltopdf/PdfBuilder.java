package test.exceltopdf.exceltopdf;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import com.itextpdf.text.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPRow;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

//import static javax.management.MBeanServerFactory.builder;

public class PdfBuilder {
	private OutputStream os;
	private Document doc;
	private PdfWriter writer;
	private List<PdfPTable> tables = new ArrayList<PdfPTable>();
	private PdfPageSetting setting = new PdfPageSetting();;
	
	public PdfBuilder(OutputStream os) {
		if(os == null){
			throw new RuntimeException("输出流os不能为空");
		}
		this.os = os;
	}
	
	public PdfBuilder buildDocument(){
		doc = new Document();
		try {
			writer = PdfWriter.getInstance(doc, os);
		} catch (DocumentException e) {
			e.printStackTrace();
		}
		return this;
	}
	
	public PdfBuilder buildTable(SheetStruct sheet,int SheetNum){
		System.out.println("SheetNum:"+SheetNum);
		int colNum = sheet.colNum()+1;
		float[] ratio = new float[colNum];
		for(int i = 0; i < ratio.length; i++){
			ratio[i]=1f/colNum;
		}
		ratio[colNum-1]=0.01f;
		PdfPTable table = new PdfPTable(ratio);
		int cellNum = colNum*sheet.rowNum();
		for(int i =0 ;i<cellNum;i++){
			PdfPCell pdfcell = new PdfPCell(ParaFactory.size8(" "));
			if(SheetNum==0||SheetNum==1||SheetNum==2){
				pdfcell.disableBorderSide(1);
				pdfcell.disableBorderSide(2);
				pdfcell.disableBorderSide(4);
				pdfcell.disableBorderSide(8);
			}

			table.addCell(pdfcell);
		}
		Iterator<Row> rows = sheet.getSheet().rowIterator();
		while(rows.hasNext()){
			Row row = rows.next();
			Iterator<Cell> cells = row.cellIterator();
			while(cells.hasNext()){
				Cell cell = cells.next();

				if(sheet.isBeMeger(cell)){
					table.getRow(cell.getRowIndex()).getCells()[cell.getColumnIndex()]=null;
					continue;
				}
				CellCoordinate index = new CellCoordinate(cell.getRowIndex(),cell.getColumnIndex());
				cell.setCellType(Cell.CELL_TYPE_STRING);
				if(sheet.getMergeCell().containsKey(index)){
					CellRangeAddress cra = sheet.getMergeCell().get(index);
					int colSpan = cra.getLastColumn()-cra.getFirstColumn()+1;
					int rowSpan = cra.getLastRow()-cra.getFirstRow()+1;
					PdfPCell pdfcell = table.getRow(cra.getFirstRow()).getCells()[cra.getFirstColumn()];
					pdfcell.setColspan(colSpan);
					pdfcell.setRowspan(rowSpan);

//					if(colSpan==9&&rowSpan==19){
//						System.out.println("asdfasdfasdf");
//						try {
//							Image img = Image.getInstance("test-file/0_1_0.png");
//							img.scaleAbsolute(300,200);
//							img.setAlignment(Image.ALIGN_LEFT|Image.TEXTWRAP);
//							img.setAbsolutePosition(0, 0);
//							doc.open();
//							this.doc.add(img);
//						} catch (Exception e) {
//							e.printStackTrace();
//						}
////						PictureUtil pu = new PictureUtil();
////						pu.createPDF3();
//					}

				}
				String val = cell.getStringCellValue();
				PdfPCell pdfcell = table.getRow(cell.getRowIndex()).getCells()[cell.getColumnIndex()];
				pdfcell.setPhrase(ParaFactory.size8(val));
				StyleSyncer.syncStyle(cell, pdfcell);
			}
		}
		for(PdfPRow row : table.getRows()){
			PdfPCell pdfcell = row.getCells()[table.getNumberOfColumns()-1];
			pdfcell.setBorder(PdfPCell.NO_BORDER);

		}
		table.setWidthPercentage(100);
		tables.add(table);
		return this;
	}
	
	public PdfBuilder buildPageSetting(float marginLeft, float marginRight,
			float marginTop, float marginBottom){
		setting.setMarginLeft(marginLeft);
		setting.setMarginRight(marginRight);
		setting.setMarginTop(marginTop);
		setting.setMarginBottom(marginBottom);
		return this;
	}
	
	public PdfBuilder buildPageSetting(Rectangle pageSize){
		setting.setPageSize(pageSize);
		return this;
	}
	
	public PdfBuilder build(){
		doc.setPageSize(setting.getPageSize());
		//将多出的一行多出的0.3毫米转化为像素
		float left = (float) (0.3*72/2.54);
		doc.setMargins(setting.getMarginLeft(), 
				setting.getMarginRight()-left, 
				setting.getMarginTop(), 
				setting.getMarginBottom());
		doc.open();
		try {
			for(PdfPTable table : tables){
				doc.add(table);
				doc.newPage();
			}
		} catch (DocumentException e) {
			e.printStackTrace();
		}
		doc.close();
		return this;
	}

	public OutputStream getOs() {
		return os;
	}

	public void setOs(OutputStream os) {
		this.os = os;
	}

	public Document getDoc() {
		return doc;
	}

	public void setDoc(Document doc) {
		this.doc = doc;
	}

	public PdfWriter getWriter() {
		return writer;
	}

	public void setWriter(PdfWriter writer) {
		this.writer = writer;
	}

	public List<PdfPTable> getTables() {
		return tables;
	}

	public void setTables(List<PdfPTable> tables) {
		this.tables = tables;
	}

	public PdfPageSetting getSetting() {
		return setting;
	}

	public void setSetting(PdfPageSetting setting) {
		this.setting = setting;
	}
	
	
}
