package test.exceltopdf.exceltopdf;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;

import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Paragraph;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import test.exceltopdf.api.ConvertFactory;

import com.itextpdf.text.PageSize;
import test.exceltopdf.test.PdfTest;

public class ExcelToPdfFactory implements ConvertFactory{
	
	public static void execute(String srcPath, OutputStream os){
		new ExcelToPdfFactory().convert(srcPath, os);
	}
	

	public void convert(String srcPath, OutputStream os) {
		Workbook wb = createdWorkBook(srcPath);
		WorkBookStruct wbStruct = new WorkBookStruct(wb);
		PdfBuilder builder = new PdfBuilder(os);
		builder.buildDocument()
		.buildPageSetting(10, 10, 10, 10)
		.buildPageSetting(PageSize.A4);
		builder.getDoc().open();
//		try {
//			builder.getDoc().add(new Paragraph("lllllllll"));
//		} catch (DocumentException e) {
//			e.printStackTrace();
//		}
		//获取excel sheet总数
		int sheetNumbers = wb.getNumberOfSheets();
		// 循环sheet
		for (int i = 0; i < sheetNumbers; i++) {
			int sheetNum = i;
			SheetStruct sheetStruct = new SheetStruct(wbStruct.getSheets()[i]);
			builder.buildTable(sheetStruct,sheetNum);
		}
//		for(Sheet sheet : wbStruct.getSheets()){
//			SheetStruct sheetStruct = new SheetStruct(sheet);
//			builder.buildTable(sheetStruct);
//		}
		builder.build();
		PdfTest pt = new PdfTest();
		pt.appendImg("test-file/test1.pdf","test-file/test3.pdf");
	}
	
	
	private Workbook createdWorkBook(String srcPath){
		File src = new File(srcPath);
		Workbook wb;
		try {
			wb = WorkbookFactory.create(src);
			return wb;
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}
}
