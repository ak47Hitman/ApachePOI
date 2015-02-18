import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by logvinov on 18.02.2015.
 */
public class CreateExcel {
    public void createExcelDoc(String inputText) throws IOException {
        Workbook wb = new HSSFWorkbook();
        ParsingHtml parsingHtml = new ParsingHtml(inputText);
        CreationHelper createHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet(parsingHtml.getTitle());

        // Create a row and put some cells in it. Rows are 0 based.

        Row row = sheet.createRow((short)0);
        // Create a cell and put a value in it.
        sheet.addMergedRegion(new CellRangeAddress(
                0, //first row (0-based)
                0, //last row  (0-based)
                0, //first column (0-based)
                parsingHtml.getNumberOfTDForMergeSells()  //last column  (0-based)
        ));

        Cell cell = row.createCell(0)
                ;
        cell.setCellValue(parsingHtml.getHeader());

//        CellUtil.setAlignment(cell, wb, CellStyle.ALIGN_CENTER);

        CellStyle cs = wb.createCellStyle();
        Font f = wb.createFont();
//        cell.setCellStyle(Font.BOLDWEIGHT_BOLD);
        f.setBoldweight(Font.BOLDWEIGHT_BOLD);

        cs.setFont(f);
        cs.setAlignment(CellStyle.ALIGN_CENTER);
        cs.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
        cs.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
        cs.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
        cs.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
//        sheet.setDefaultColumnStyle(0, cs); //set bold for column 1
//        row.getCell(0).setCellStyle(cs);
        cell.setCellStyle(cs);



        // Or do it on one line.
        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue(
                createHelper.createRichTextString("This is a string"));
        row.createCell(3).setCellValue(true);
        Row row2 = sheet.createRow((short)1);
        Cell cell2 = row2.createCell(0);
        cell2.setCellValue(1);

        // Or do it on one line.
        row2.createCell(1).setCellValue("21/11/14");
        row2.createCell(2).setCellValue(
                createHelper.createRichTextString("This is a string"));
        row2.createCell(3).setCellValue(true);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(parsingHtml.getTitle() + ".xls");
        wb.write(fileOut);
        fileOut.close();
    }
}
