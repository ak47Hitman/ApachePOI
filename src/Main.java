import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.*;

import java.util.Calendar;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {

    public static void main(String[] args) throws IOException {
//        String testFileName = "C:\\Users\\logvinov\\IdeaProjects\\ApachePOI\\1.html";
        String testFileName = "C:\\Users\\logvinov\\Desktop\\MyWorkSpace\\ApachePOI\\1.xls";

        TextFile textFile = new TextFile();
        CreateExcel createExcel = new CreateExcel();
        String text1 = textFile.read(testFileName);
        createExcel.createExcelDoc(text1);
//        ParsingHtml parsingHtml = new ParsingHtml(text1);
//        parsingHtml.getTitle();
//        parsingHtml.getHeader();
//        parsingHtml.getTable();

//        System.out.println(textFile.read(testFileName));
//        if ((new File(outPutFileInfo)).exists()) {
//            // существует
//        } else {
//            // не существует
//            textFile.write(outPutFileInfo, "");
//        }

//        //Получение атрибутов тега <A>
//        HTMLEditorKit kit = new HTMLEditorKit();
//        URL ura = new URL("http://vingrad.ru");
//        try {
//            InputStream in = ura.openStream();
//            doc = (HTMLDocument) kit.createDefaultDocument();
//            doc.putProperty("IgnoreCharsetDirective",Boolean.TRUE);
//            kit.read(in, doc, 0);
//        } catch (BadLocationException ex) {
//            ex.printStackTrace();
//        } catch (IOException ex) {
//            ex.printStackTrace();
//        }
//        HTMLDocument.Iterator it = doc.getIterator(HTML.Tag.A);
//        while (it.isValid()) {
//            AttributeSet attrs = it.getAttributes();
//            Object linkAttr = attrs.getAttribute(HTML.Attribute.HREF);
//            System.out.println(linkAttr);
//            it.next();
//        }
//        Parser parser = new Parser ("http://whatever");
//        NodeList list = parser.parse (null);
//        Node node = list.elementAt(0);
//        NodeList sublist = node.getChildren ();
//        System.out.println (sublist.size ());
        // do something with your list of nodes.


//        parseWithJsoup(text1);
//        System.out.println("Prev 0 : " + text1);
//
//        Pattern pattern = Pattern.compile("<td>(.*?)</td>");
////        Pattern pattern = Pattern.compile("<tr>(.*?)</tr>");
////        Pattern pattern = Pattern.compile("(?i)<title([^>]+)>(.+?)</title>");
//        Matcher matcher = pattern.matcher(text1);
//        System.out.println("Text1: " + text1);
//        while(matcher.find()) {
//            System.out.println("Find text: " + matcher.group(1));
////            prevArray.add(matcher.group(1));
//        }
//
//        Workbook wb = new HSSFWorkbook();
//        //Workbook wb = new XSSFWorkbook();
//        CreationHelper createHelper = wb.getCreationHelper();
//        Sheet sheet = wb.createSheet("new sheet");
//
//        // Create a row and put some cells in it. Rows are 0 based.
//        Row row = sheet.createRow(0);
//
//        // Create a cell and put a date value in it.  The first cell is not styled
//        // as a date.
//        Cell cell = row.createCell(0);
//        cell.setCellValue(new Date());
//
//        // we style the second cell as a date (and time).  It is important to
//        // create a new cell style from the workbook otherwise you can end up
//        // modifying the built in style and effecting not only this cell but other cells.
//        CellStyle cellStyle = wb.createCellStyle();
//        cellStyle.setDataFormat(
//                createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
//        cell = row.createCell(1);
//        cell.setCellValue(new Date());
//        cell.setCellStyle(cellStyle);
//
//        //you can also set date as java.util.Calendar
//        cell = row.createCell(2);
//        cell.setCellValue(Calendar.getInstance());
//        cell.setCellStyle(cellStyle);
//
//        // Write the output to a file
//        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
//        wb.write(fileOut);
//        fileOut.close();
////        Workbook wb = new HSSFWorkbook();
////        //Workbook wb = new XSSFWorkbook();
////        CreationHelper createHelper = wb.getCreationHelper();
////        Sheet sheet = wb.createSheet("new sheet");
////
////        // Create a row and put some cells in it. Rows are 0 based.
////        Row row = sheet.createRow((short)0);
////        // Create a cell and put a value in it.
////        Cell cell = row.createCell(0);
////        cell.setCellValue(1);
////
////        // Or do it on one line.
////        row.createCell(1).setCellValue(1.2);
////        row.createCell(2).setCellValue(
////                createHelper.createRichTextString("This is a string"));
////        row.createCell(3).setCellValue(true);
////
////        // Write the output to a file
////        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
////        wb.write(fileOut);
////        fileOut.close();
//////        Workbook wb = new HSSFWorkbook();  // or new XSSFWorkbook();
//////        Sheet sheet1 = wb.createSheet("new sheet");
//////        Sheet sheet2 = wb.createSheet("second sheet");
//////
//////        // Note that sheet name is Excel must not exceed 31 characters
//////        // and must not contain any of the any of the following characters:
//////        // 0x0000
//////        // 0x0003
//////        // colon (:)
//////        // backslash (\)
//////        // asterisk (*)
//////        // question mark (?)
//////        // forward slash (/)
//////        // opening square bracket ([)
//////        // closing square bracket (])
//////
//////        // You can use org.apache.poi.ss.util.WorkbookUtil#createSafeSheetName(String nameProposal)}
//////        // for a safe way to create valid names, this utility replaces invalid characters with a space (' ')
//////        String safeName = WorkbookUtil.createSafeSheetName("[O'Brien's sales*?]"); // returns " O'Brien's sales   "
//////        Sheet sheet3 = wb.createSheet(safeName);
//////
//////        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
//////        wb.write(fileOut);
//////        fileOut.close();
////////        Workbook wb = new HSSFWorkbook();
////////        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
////////        wb.write(fileOut);
////////        fileOut.close();
        System.out.println("All done");
    }
    private static void parseWithJsoup(String HTML) {
        System.out.println("*** JSOUP ***");
        Document doc = Jsoup.parse(HTML);
        System.out.println("Title: " + doc.getElementsByTag("title").text());
        System.out.println("H1: " + doc.getElementsByTag("h1").text());
        Element table = doc.getElementsByTag("table").first();
        Elements trs = table.getElementsByTag("tr");
        for (Element tr : trs) {
            System.out.println("TR: " + tr.text());
            for (Element td : tr.getAllElements()) {
                System.out.println("TD: " + td.text());
            }
        }
        System.out.println();
    }
}
