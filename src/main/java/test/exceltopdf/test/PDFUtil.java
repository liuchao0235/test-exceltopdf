package test.exceltopdf.test;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.log.Logger;
import com.itextpdf.text.log.LoggerFactory;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.XMLWorkerFontProvider;
import com.itextpdf.tool.xml.XMLWorkerHelper;

import java.io.*;
import java.nio.charset.Charset;

public class PDFUtil {
    private static final Logger logger = LoggerFactory.getLogger(PDFUtil.class);
    private static String fontPath = "test-file/simsun.ttf";
    public static InputStream htmlToPDF(InputStream htmlInputStream) {
        ByteArrayOutputStream out = null;
        ByteArrayInputStream inputStream = null;
        Document document = new Document();
        XMLWorkerFontProvider provider = new XMLWorkerFontProvider();
        provider.register(fontPath);
        try {
            out = new ByteArrayOutputStream();
            PdfWriter writer = PdfWriter.getInstance(document, out);
            document.open();
            XMLWorkerHelper.getInstance().parseXHtml(writer, document, new BufferedInputStream(htmlInputStream));
//            XMLWorkerHelper.getInstance().parseXHtml(writer, document,htmlInputStream);
            document.close();
            inputStream = new ByteArrayInputStream(out.toByteArray());
        } catch (DocumentException e) {
            logger.error(e.getMessage(), e);
        } catch (IOException e) {
            document.close();
            logger.error(e.getMessage(), e);
        }
        finally {
            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    logger.error(e.getMessage(), e);
                }
            }
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    logger.error(e.getMessage(), e);
                }
            }
        }        return inputStream;
    }
    public static void main(String[] args) throws IOException{
        String html = "C:/Users/rose/Desktop/2018-02-11/fed828eb-aff4-49a4-81e2-5c5ed871c63b.2018.02.08.html";
//        String cssInput = "C:/Users/rose/Desktop/2018-02-11/445b4964-4a7b-48d9-b6d3-9385781c5467.2018.02.08.files/stylesheet.css";
        String pdf = "test-file/htmltopdf.pdf";
        File pdfFile = new File(pdf);
        InputStream inputStream = htmlToPDF(new FileInputStream(new File(html)));
        try{
            if(!pdfFile.exists()){
                pdfFile.createNewFile();
            }
        }catch (IOException e ){
            //
        }
        IOUtil.copyCompletely(inputStream,new FileOutputStream(pdf));
    }
}