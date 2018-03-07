package test.exceltopdf.test;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.PdfWriter;
import java.io.File;
import com.itextpdf.tool.xml.XMLWorkerHelper;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
/**
  *
  * @author iText
  */
public class ParseHtmlTable {
public static final String DEST = "test-file/html_table.pdf" ;
public static final String HTML = "test-file/table_css.html";

        public static void main(String[] args) throws IOException, DocumentException {

    File file = new File( DEST);
             file.getParentFile().mkdirs();
    //new ParseHtmlTable5 ().createPdf(DEST );
}
       /**
      * Creates a PDF with the words "Hello World"
      * @param file
      * @throws IOException
      * @throws DocumentException
      */
        public void createPdf(String file) throws IOException, DocumentException {
            Document document = new Document();
             PdfWriter writer = PdfWriter. getInstance(document, new FileOutputStream(file));
// step 3
document.open();
// step 4
XMLWorkerHelper. getInstance().parseXHtml(writer, document,
                 new FileInputStream( HTML));
document.close();
 }
}
