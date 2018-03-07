package test.exceltopdf.exceltopdf;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import java.io.*;
import java.net.MalformedURLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class PictureUtil {
    public static Map<String, PictureData> getSheetPictrues03(int sheetNum,
                                                              HSSFSheet sheet, HSSFWorkbook workbook) {
        Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
        List<HSSFPictureData> pictures = workbook.getAllPictures();
        if (pictures.size() != 0) {
            for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
                HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
                if (shape instanceof HSSFPicture) {
                    HSSFPicture pic = (HSSFPicture) shape;
                    int pictureIndex = pic.getPictureIndex() - 1;
                    HSSFPictureData picData = pictures.get(pictureIndex);
                    String picIndex = String.valueOf(sheetNum) + "_"
                            + String.valueOf(anchor.getRow1()) + "_"
                            + String.valueOf(anchor.getCol1());
                    sheetIndexPicMap.put(picIndex, picData);
                }
            }
            return sheetIndexPicMap;
        } else {
            return null;
        }
    }

    //07格式excel获取图片。
    public static Map<String, PictureData> getSheetPictrues07(int sheetNum,
                                                              XSSFSheet sheet, XSSFWorkbook workbook) {
        Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();

        for (POIXMLDocumentPart dr : sheet.getRelations()) {
            if (dr instanceof XSSFDrawing) {
                XSSFDrawing drawing = (XSSFDrawing) dr;
                List<XSSFShape> shapes = drawing.getShapes();
                for (XSSFShape shape : shapes) {
                    XSSFPicture pic = (XSSFPicture) shape;
                    XSSFClientAnchor anchor = pic.getPreferredSize();
                    CTMarker ctMarker = anchor.getFrom();
                    String picIndex = String.valueOf(sheetNum) + "_"
                            + ctMarker.getRow() + "_" + ctMarker.getCol();
//                    pic.getPictureData().getData();
                    sheetIndexPicMap.put(picIndex, pic.getPictureData());
                }
            }
        }

        return sheetIndexPicMap;
    }

//图片及位置获取

    public static void getAllImg(String excelPath) throws InvalidFormatException, IOException {

        // 创建文件
        File file = new File(excelPath);

        // 创建流
        InputStream input = new FileInputStream(file);

        // 获取文件后缀名
        String fileExt =  file.getName().substring(file.getName().lastIndexOf(".") + 1);

        // 创建Workbook
        Workbook wb = null;

        // 创建sheet
        Sheet sheet = null;

        //根据后缀判断excel 2003 or 2007+
        if (fileExt.equals("xls")) {
            wb = (HSSFWorkbook) WorkbookFactory.create(input);
        } else {
            wb = new XSSFWorkbook(input);
        }

        //获取excel sheet总数
        int sheetNumbers = wb.getNumberOfSheets();

        // sheet list
        List<Map<String, PictureData>> sheetList = new ArrayList<Map<String, PictureData>>();

        // 循环sheet
        for (int i = 0; i < sheetNumbers; i++) {

            sheet = wb.getSheetAt(i);
            // map等待存储excel图片
            Map<String, PictureData> sheetIndexPicMap;

            // 判断用07还是03的方法获取图片
            if (fileExt.equals("xls")) {
                sheetIndexPicMap = getSheetPictrues03(i, (HSSFSheet) sheet, (HSSFWorkbook) wb);
            } else {
                sheetIndexPicMap = getSheetPictrues07(i, (XSSFSheet) sheet, (XSSFWorkbook) wb);
            }
            // 将当前sheet图片map存入list
            sheetList.add(sheetIndexPicMap);
        }
//        Map map = getData(excelPath);
        printImg(sheetList);

    }

    //将图片保存到指定位置
        public static void printImg(List<Map<String, PictureData>> sheetList) throws IOException {

        for (Map<String, PictureData> map : sheetList) {
            Object key[] = map.keySet().toArray();
            for (int i = 0; i < map.size(); i++) {
                // 获取图片流
                PictureData pic = map.get(key[i]);
                // 获取图片索引
                String picName = key[i].toString();
                // 获取图片格式
                String ext = pic.suggestFileExtension();

                byte[] data = pic.getData();

                FileOutputStream out = new FileOutputStream("test-file/" + picName + "." + ext);
                out.write(data);
                out.close();
            }

        }

    }

    public void createPDF()
    {
        try
        {
            String RESULT = "F:\\java56班\\eclipse-SDK-4.2-win32\\pdfiText.pdf";
            Document document = new Document();
            PdfWriter writer = PdfWriter.getInstance(document,
                    new FileOutputStream(RESULT));
            document.open();
            PdfContentByte canvas = writer.getDirectContentUnder();
            writer.setCompressionLevel(0);
            canvas.saveState(); // q
            canvas.beginText(); // BT
            canvas.moveText(36, 788); // 36 788 Td
            canvas.setFontAndSize(BaseFont.createFont(), 12); // /F1 12 Tf
            // canvas.showText("Hello World"); // (Hello World)Tj
            canvas.showText("你好"); // (Hello World)Tj
            canvas.endText(); // ET
            canvas.restoreState(); // Q
            document.close();
        } catch (FileNotFoundException e)
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (DocumentException e)
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e)
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    public void createPDF2()
    {
        try
        {
            String RESULT = "F:\\java56班\\eclipse-SDK-4.2-win32\\pdfiText.pdf";
            Document document = new Document();
            PdfWriter writer = PdfWriter.getInstance(document,
                    new FileOutputStream(RESULT));
            document.open();
            writer.setCompressionLevel(0);
            // Phrase hello = new Phrase("Hello World");
            Phrase hello = new Phrase("你好");
            PdfContentByte canvas = writer.getDirectContentUnder();
            ColumnText.showTextAligned(canvas, Element.ALIGN_LEFT, hello, 200,
                    800, 0);
            document.close();
        } catch (FileNotFoundException e)
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (DocumentException e)
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    public void createPDF3()
    {
        try
        {
            String resource_jpg = "test-file/0_1_0.png";//
            String RESULT = "test-file/test2.pdf";
//            Paragraph p = new Paragraph("Foobar Film Festival", new Font(Font.FontFamily.HELVETICA, 24));
//            p.setAlignment(Element.ALIGN_CENTER);
            Document document = new Document();//
            PdfWriter writer = PdfWriter.getInstance(document,new FileOutputStream(RESULT));
            document.open();
//            document.add(p);
            Image img = Image.getInstance(resource_jpg);
            img.scaleAbsolute(300,200);
            img.setAlignment(Image.ALIGN_LEFT|Image.TEXTWRAP);
//            img.setAbsolutePosition(
//                    (PageSize.POSTCARD.getWidth() - img.getScaledWidth()) / 2,
//                    (PageSize.POSTCARD.getHeight() - img.getScaledHeight()) / 2);

            img.setAbsolutePosition(0, 0);
            document.add(img);
//            document.newPage();
//            document.add(p);
//            document.add(img);
//            PdfContentByte over = writer.getDirectContent();
//            over.saveState();
//            float sinus = (float) Math.sin(Math.PI / 60);
//            float cosinus = (float) Math.cos(Math.PI / 60);
//            BaseFont bf = BaseFont.createFont();
//            over.beginText();
//            over.setTextRenderingMode(PdfContentByte.TEXT_RENDER_MODE_FILL_STROKE);
//            over.setLineWidth(1.5f);
//            over.setRGBColorStroke(0xFF, 0x00, 0x00);
//            over.setRGBColorFill(0xFF, 0xFF, 0xFF);
//            over.setFontAndSize(bf, 36);
//            over.setTextMatrix(cosinus, sinus, -sinus, cosinus, 50, 324);
//            over.showText("SOLD OUT");
//            over.endText();
//            over.restoreState();
//            PdfContentByte under = writer.getDirectContentUnder();
//            under.saveState();
//            under.setRGBColorFill(0xFF, 0xD7, 0x00);
//            under.rectangle(5, 5, PageSize.POSTCARD.getWidth() - 10,
//                    PageSize.POSTCARD.getHeight() - 10);
//            under.fill();
//            under.restoreState();

            document.close();
        } catch (FileNotFoundException e)
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (DocumentException e)
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (MalformedURLException e)
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e)
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
    public static void main(String[] args) throws Exception{
        PictureUtil pu = new PictureUtil();
        String excelPath = "test-file/test1.xlsx";
        pu.getAllImg(excelPath);
//        pu.createPDF3();
    }
}
