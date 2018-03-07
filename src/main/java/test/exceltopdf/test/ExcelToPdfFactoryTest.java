package test.exceltopdf.test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import test.exceltopdf.exceltopdf.ExcelToPdfFactory;
import test.exceltopdf.exceltopdf.PictureUtil;

public class ExcelToPdfFactoryTest {

	public static void main(String args[]) {
		OutputStream os = null;
		try {
			os = new FileOutputStream("test-file/test1.pdf");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		PictureUtil pu = new PictureUtil();
		String excelPath = "test-file/test1.xlsx";
		try {
			pu.getAllImg(excelPath);
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		ExcelToPdfFactory.execute(excelPath, os);
	}

}
