import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by logvinov on 18.02.2015.
 */
public class CreateExcel {
    public void createExcelDoc(String inputText) throws IOException {
        // Задаем два индекса для перебора ячеек и строк. При создании ячейки или строки он увеличивается на 1 или обнуляется - если строка закончена.
        int numberOfRows = 0;
        int numberOfCells = 0;
        // Создаем книгу
        Workbook wb = new HSSFWorkbook();
        // Отдаем текст для парсинга текста из входного файла
        ParsingHtml parsingHtml = new ParsingHtml(inputText);
        // Создаем лист. Название берем из заголовка
        Sheet sheet = wb.createSheet(parsingHtml.getTitle());
        // Создаем первую строку в таблице - она будет заголовком.
        Row row = sheet.createRow((short)numberOfRows);
        // Объединяем регион для вставки заголвока. Для определения на сколько ячеек нужно обяединить строку мы высчитываем максимально
        // возможное число ячеек в строках таблицы
        sheet.addMergedRegion(new CellRangeAddress(
                numberOfRows, // указываем начальную строку в которой объединяем ячейки
                numberOfRows, // указываем конечную строку в которой объединяем ячейки (у нас начальная и конечная строки совпадают)
                numberOfCells, // указываем начальную ячейку - унас она равна 0 - потомучто объеждиняем с первой ячейки
                parsingHtml.getNumberOfTDForMergeCells() - 1  // указываем последюю ячейку для обеъединения - она у нас
                                                              // равна максимальной возможной ячейке из всех строк таблицы
        ));
        // Создаем ячейку в первой строке
        Cell cell = row.createCell(numberOfCells);
        // Укладываем в нее наш текст из заголовка
        cell.setCellValue(parsingHtml.getHeader());
        // Заголовок должен быть жирным, а в тегах ничего про Жирность не указывается - поэтому мы задаем жирность РУКАМИ
        Font fontBOLD = wb.createFont();
        fontBOLD.setBoldweight(Font.BOLDWEIGHT_BOLD);
        // Не жирный тип. Используется для смены жирного на нежирный
        Font fontNormal = wb.createFont();
        fontNormal.setBoldweight(Font.BOLDWEIGHT_NORMAL);
        // Задаем стиль для заголовка.
        CellStyle headerStyle = wb.createCellStyle();
        // Говорим что текст должен быть жирным
        headerStyle.setFont(fontBOLD);
        // А так же указываем что заголовок должен быть по центру
        headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
        // Укладываем стиль
        cell.setCellStyle(headerStyle);
        // Так как ячейка заполнена, и строка тоже - увеличиваем нашу строку на 1.
        numberOfRows++;
        //
        row = sheet.createRow((short)numberOfRows);
        sheet.addMergedRegion(new CellRangeAddress(
                numberOfRows, //first row (0-based)
                numberOfRows, //last row  (0-based)
                0, //first column (0-based)
                parsingHtml.getNumberOfTDForMergeCells() - 1  //last column  (0-based)
        ));
        Cell cell2 = row.createCell(numberOfCells);
        cell2.setCellValue("");
        numberOfRows++;

        System.out.println("Cells: " + numberOfCells + " Rows :" + numberOfRows);
        CellStyle cs = wb.createCellStyle();
        cs.setAlignment(CellStyle.ALIGN_CENTER);

        CellStyle standartStyle = wb.createCellStyle();
        standartStyle.setAlignment(CellStyle.ALIGN_CENTER);
        standartStyle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
        standartStyle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
        standartStyle.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
        standartStyle.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
        standartStyle.setWrapText(true);

        CellStyle dateType = wb.createCellStyle();
        dateType.setAlignment(CellStyle.ALIGN_CENTER);
        dateType.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
        dateType.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
        dateType.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
        dateType.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
        dateType.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        dateType.setFillPattern(headerStyle.SOLID_FOREGROUND);
//        dateType.setDataFormat(wb.createDataFormat().getFormat("m/d/yy"));
        dateType.setWrapText(true);

        CellStyle rowStyleWithForeGround = wb.createCellStyle();
        rowStyleWithForeGround.setAlignment(CellStyle.ALIGN_CENTER);
        rowStyleWithForeGround.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
        rowStyleWithForeGround.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
        rowStyleWithForeGround.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
        rowStyleWithForeGround.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
        rowStyleWithForeGround.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        rowStyleWithForeGround.setFillPattern(headerStyle.SOLID_FOREGROUND);
        rowStyleWithForeGround.setWrapText(true);

        CellStyle rowStyleBold = wb.createCellStyle();
        rowStyleBold.setAlignment(CellStyle.ALIGN_CENTER);
        rowStyleBold.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
        rowStyleBold.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
        rowStyleBold.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
        rowStyleBold.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
        rowStyleBold.setFont(fontBOLD);
        rowStyleBold.setWrapText(true);

        CellStyle rowStyleBoldAndForeGround = wb.createCellStyle();
        rowStyleBoldAndForeGround.setAlignment(CellStyle.ALIGN_CENTER);
        rowStyleBoldAndForeGround.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
        rowStyleBoldAndForeGround.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
        rowStyleBoldAndForeGround.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
        rowStyleBoldAndForeGround.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
        rowStyleBoldAndForeGround.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        rowStyleBoldAndForeGround.setFillPattern(headerStyle.SOLID_FOREGROUND);
        rowStyleBoldAndForeGround.setFont(fontBOLD);
        rowStyleBoldAndForeGround.setWrapText(true);

        Document doc = Jsoup.parse(inputText);
        Elements tables = doc.getElementsByTag("table");
        int numberOfTable = 1;
        for (Element table : tables) {
            Elements trs = table.getElementsByTag("tr");
            for (Element tr : trs) {
                Row rowInTable = sheet.createRow(numberOfRows);
                System.out.println("Cells: " + numberOfCells + " Rows :" + numberOfRows);
                System.out.println("TR: " + tr.text() + "Attr: " + tr.attr("style"));
                numberOfCells = 0;
                int i = 0;
                //Получаем все значения между тегами <TD>. Внутри могут находиться теги <B>. Поэтоу их тоже не забыть обработать.
                // Так же есть входная дата, которая отличается маской от Екселевской. Нужно это задать.
                //Так же разбиваем наш файл на две таблицы - первая таблица это у нас таблица заголовка, вторая таблица - остальная информация.
                // Парсинг этих строк отличается
                for (Element tds : tr.getElementsByTag("td")) {
                    if (numberOfTable == 1) {
                        //Если у нас на входе первая таблица, то делаем проверку и делаем выборку по первой таблице, в которой не проверяем на жирность заголовков
                        // и по первой таблице устанавливаем ширину таблицы по умолчанию равной ширине текста в ячейке
                        System.out.println("TD1: " + tds.text());
                            Cell cellInTable = rowInTable.createCell(numberOfCells);
                            cellInTable.setCellValue(tds.text());
                            cellInTable.setCellStyle(cs);
                            numberOfCells++;
                        } else {
                        //Если у нас вторая таблица - то тут применяем жестокий парсинг ячеек.
                        // Для начала начинаем разбор ячейки на присутствие в ней даты - если дата есть то парсим строку и находим в ней нужные символы для замены / на точку .
                        // А так же задаем тип ячейки - дата и указываем маску для ячейки
                        System.out.println("TD2: " + tds.text());
                        String findDateString;
                        //Задаем шаблон для парсинга даты через регулярку
                        Pattern pattern = Pattern.compile("(0?[1-9]|[12][0-9]|3[01])/(0?[1-9]|1[012])/(\\d\\d)");
                        //Ищем совпадение в строке
                        Matcher matcher = pattern.matcher(tds.text());
                        //Перебираем все тэги <B> чтобы добавить жирный текст
                        Elements bTag = tds.getElementsByTag("b");
                        //Создаем ячейку в строке
                        Cell cellInTable = rowInTable.createCell(numberOfCells);
                        //Задаем начало и конец жирного текста
                        int indexOfFirstElement = tds.text().indexOf(bTag.text());
                        int indexOfLastElement = bTag.text().length();
                        RichTextString richString;
                        //Если нашли дату, то начинаем применять парсинг для даты - делаем нужный нам формат из "##/##/##" в "##.##.20##" и устанавливаем тип текста если дата жирная
                        if(matcher.find()) {
                            //конвертирование нашей строки и внесение ее в ячейку
                            findDateString = tds.text().replace("/" + tds.text().substring(tds.text().length() - 2, tds.text().length()), "/20" + tds.text().substring(tds.text().length() - 2 ,tds.text().length())).replace("/", ".");
                            richString = new HSSFRichTextString(findDateString);
                            if (bTag.text().length() != 0) {
                                richString.applyFont(indexOfFirstElement, indexOfLastElement + 2, fontBOLD );
                                richString.applyFont(indexOfLastElement + 2, findDateString.length(), fontNormal );
                            }
                            //в зависимости от жирности установка нужного стиля ячейки
                            System.out.println("Length: bTag.text().length()" + bTag.text().length());
                            if (tr.attr("style").length() != 0) {
                                System.out.println("Length: bTag.text().length()" + bTag.text().length());
                                cellInTable.setCellStyle(dateType);
                            } else {
                                cellInTable.setCellStyle(standartStyle);
                            }
                        }
                        else {
                            richString = new HSSFRichTextString(tds.text());
                            if (bTag.text().length() != 0) {
                                richString.applyFont( indexOfFirstElement, indexOfLastElement, fontBOLD );
                                richString.applyFont(indexOfLastElement, tds.text().length(), fontNormal );
                                if (tr.attr("style").length() != 0) {
                                    cellInTable.setCellStyle(rowStyleWithForeGround);
                                } else {
                                    cellInTable.setCellStyle(standartStyle);
                                }
                            } else {
                                if (tr.attr("style").length() != 0) {
                                    cellInTable.setCellStyle(rowStyleWithForeGround);
                                } else {
                                    cellInTable.setCellStyle(standartStyle);
                                }
                            }
                        }
                        if (sheet.getRow(2).getCell(i).getStringCellValue().length() > tds.text().length()) {
                            sheet.setColumnWidth(numberOfCells, sheet.getRow(2).getCell(numberOfCells).getStringCellValue().length() * 256);
                        }
                        cellInTable.setCellValue(richString);
                        numberOfCells++;
                        }
                }
                numberOfTable++;
                numberOfRows++;
            }
        }
        HSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
        for (int i = 2; i < sheet.getLastRowNum(); i++)
        {
            for (int k = 0; k < sheet.getRow(i).getLastCellNum(); k++) {
                System.out.println("Last cell Number: " + sheet.getRow(i).getLastCellNum());
                System.out.println("String Cell Value in Head: " + sheet.getRow(2).getCell(k).getStringCellValue());
                System.out.println("String Cell Value not in Head: " + sheet.getRow(i).getCell(k).getStringCellValue());
                if (sheet.getRow(i).getCell(k).getStringCellValue() != null && sheet.getRow(i).getCell(k).getStringCellValue().length() > sheet.getRow(2).getCell(k).getStringCellValue().length()) {
                    System.out.println(k + " i " + i);
                    sheet.setColumnWidth(k, sheet.getRow(2).getCell(k).getStringCellValue().length() * 256 * 3);
                    System.out.println(k + " i " + i);
                    System.out.println(sheet.getLastRowNum());
                }
            }
        }
        //Запись выходного файла. Название берем из заголовка
        FileOutputStream fileOut = new FileOutputStream(parsingHtml.getTitle() + ".xls");
        wb.write(fileOut);
        fileOut.close();
    }
}