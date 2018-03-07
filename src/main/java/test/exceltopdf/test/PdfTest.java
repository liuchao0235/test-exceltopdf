package test.exceltopdf.test;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.*;

import java.io.FileOutputStream;
import java.io.IOException;


public class PdfTest {
    public void appendImg(String readPdfPath, String targetPdfPath){
        PdfReader reader = null;
        try {
            reader = new PdfReader("readPdfPath");
        } catch (IOException e) {
            e.printStackTrace();
        }
        // 书签名
        String fieldName = "报告说明";
        int n = reader.getNumberOfPages();
        Document document = new Document(reader.getPageSize(n));
        float width = document.getPageSize().getWidth();
        // Create a stamper that will copy the document to a new file
        PdfStamper stamper = null;
        try {
            stamper = new PdfStamper(reader,new FileOutputStream("targetPdfPath"));
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        // 提取pdf中的表单
        AcroFields form = stamper.getAcroFields();
        PdfContentByte over;
        Image img = null;
        try {
            form.addSubstitutionFont(BaseFont.createFont("STSong-Light","UniGB-UCS2-H", BaseFont.NOT_EMBEDDED));
            // 通过域名获取所在页和坐标，左下角为起点
        int pageNo = form.getFieldPositions(fieldName).get(0).page;
        Rectangle signRect = form.getFieldPositions(fieldName).get(0).position;
        float x = signRect.getLeft();
        float y = signRect.getBottom();

            img = Image.getInstance("test-file/4_19_0.png");
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

//        img.setAbsolutePosition(x,y);
//        img.setAlignment(Image.ALIGN_RIGHT);
        img.scaleAbsolute(300,110);
        img.setAbsolutePosition(20, 340);
        img.setAlignment(com.itextpdf.text.Image.ALIGN_LEFT| com.itextpdf.text.Image.TEXTWRAP);
        width = width - img.getWidth();
        System.out.println("width:" + width);
        if(n > 0){
            // Text over the existing page
            over = stamper.getOverContent(9);
            try {
                over.addImage(img);
            } catch (DocumentException e) {
                e.printStackTrace();
            }
        }
        try {
            stamper.close();
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        reader.close();
    }
    public static void main(String[] args) throws Exception {
            PdfTest pt = new PdfTest();
            pt.appendImg("test-file/test1.pdf","test-file/test3.pdf");


    }
}
