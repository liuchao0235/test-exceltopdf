package test.exceltopdf.exceltopdf;

import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.Paragraph;

public class ParaFactory {
	private static Font defFont(){
		return FontFactory.getFont("STSong-Light", "UniGB-UCS2-H");
	}
	public static Paragraph size10(String content){
		Paragraph p = new Paragraph(content, defFont());
		p.getFont().setSize(10f);
		return p;
	}
	public static Paragraph size8(String content){
		Paragraph p = new Paragraph(content, defFont());
		p.getFont().setSize(8f);
		return p;
	}
}
