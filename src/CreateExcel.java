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
        Row headerFirstRow = sheet.createRow((short)numberOfRows);
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
        Cell headerCell = headerFirstRow.createCell(numberOfCells);
        // Укладываем в нее наш текст из заголовка
        headerCell.setCellValue(parsingHtml.getHeader());
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
        headerStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // Укладываем стиль
        headerCell.setCellStyle(headerStyle);
        // Так как ячейка заполнена, и строка тоже - увеличиваем нашу строку на 1.
        numberOfRows++;
        // Объединяем ячейки и всталяем пустое значение. Делается для красоты - отступа между заголовком и таблицей на вывод
        headerFirstRow = sheet.createRow((short)numberOfRows);
        // Объединяем ячейки
        sheet.addMergedRegion(new CellRangeAddress(
                numberOfRows, //first row (0-based)
                numberOfRows, //last row  (0-based)
                0, //first column (0-based)
                parsingHtml.getNumberOfTDForMergeCells() - 1  //last column  (0-based)
        ));
        // Задаем пустую строку
        Cell nullCell = headerFirstRow.createCell(numberOfCells);
        nullCell.setCellValue("");
        numberOfRows++;

        // Задаем стиль для шапки таблицы
        CellStyle headStyle = wb.createCellStyle();
        // Центрирование текста в заголовке
        headStyle.setAlignment(CellStyle.ALIGN_CENTER);
        // Задаем стиль для обычных ячеек в таблице
        CellStyle standartStyle = wb.createCellStyle();
        // Центрирование текста в заголвоке
        standartStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // Границы ячеек в таблице
        standartStyle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
        standartStyle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
        standartStyle.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
        standartStyle.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
        // Перенос текста на новую строку
        standartStyle.setWrapText(true);

        // Задаем стиль для подкрашенных ячеек таблицы
        CellStyle rowStyleWithForeGround = wb.createCellStyle();
        rowStyleWithForeGround.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        rowStyleWithForeGround.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
        rowStyleWithForeGround.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
        rowStyleWithForeGround.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
        rowStyleWithForeGround.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
        rowStyleWithForeGround.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        rowStyleWithForeGround.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        rowStyleWithForeGround.setWrapText(true);

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
                        cellInTable.setCellStyle(headStyle);
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
                            // Увеличиваем индекс на 2 добавленных элемента в дате "/20"
                            indexOfLastElement += 2;
                        }
                        // Если дату не нашли то не парсим дату и задаем форматирование для остального текста
                        else {
                            richString = new HSSFRichTextString(tds.text());
                        }
                        // Если присутствует тэг <b> в тексте
                        if (bTag.text().length() != 0) {
                            richString.applyFont(indexOfFirstElement, indexOfLastElement, fontBOLD);
                        }
                        //в зависимости от жирности установка нужного стиля ячейки
                        if (tr.attr("style").length() != 0) {
                            cellInTable.setCellStyle(rowStyleWithForeGround);
                        } else {
                            cellInTable.setCellStyle(standartStyle);
                        }
                        // Вставляем значение в ячейку
                        cellInTable.setCellValue(richString);
                        // Увеличиваем счетчик ячеек
                        numberOfCells++;
                        }
                }
                // Если есть еще таблицы в HTML - файле то увеличиваем счетчик таблиц и повторяем цикл
                numberOfTable++;
                // У величиваем счетчик строк
                numberOfRows++;
            }
        }
        // Установка ширины ячейки исходя из длинны заголовка. Перебираем каждую строку в таблице и в каждой строке перебираем ячейки.
        // Если ширина заголовка больше или равна длинне значения в ячейке, то шириной ячейки устанавливаем длинну заголовка * 256 (системное значение).
        // Если же ширина заголовка меньше длинны значения в ячейке, то устанавливаем ширину ячейки как длинна заголовка * 256 (системное значение) * 3.
        // Умножаем на 3 для красоты. При умножении на 2 слишком сжатый вид. При желании можно установить своё значение - в зависимотси от пожелания.
        // Пробегаем по всем строкам в таблице
        for (int i = 2; i < sheet.getLastRowNum(); i++)
        {   // В каждой строке пробегаем по всем ячейкам таблицы
            for (int k = 0; k < sheet.getRow(i).getLastCellNum(); k++) {
                // Если ширина колонки меньше чем ширина заголовка * на 256 (системный коэффициент) * 3, то изменяем значение.
                // Если же ширина колонки больше либо равна ширине заголовка, то это говорит о том что мы изменили значение в ячейке и такие ячейки не надо изменять.
                if(sheet.getColumnWidth(k) < sheet.getRow(2).getCell(k).getStringCellValue().length() * 256 * 3) {
                    // Если значение в ячейке не null (избежать null pointer Exception) и длинна текста в ячейке больше чем длина заголовка,
                    // то установим значение ширины ячейки как ширина заголовка * 256 * 3
                    if (sheet.getRow(i).getCell(k).getStringCellValue() != null &&
                            sheet.getRow(i).getCell(k).getStringCellValue().length() > sheet.getRow(2).getCell(k).getStringCellValue().length()) {
                        sheet.setColumnWidth(k, sheet.getRow(2).getCell(k).getStringCellValue().length() * 256 * 3);
                        System.out.println(sheet.getLastRowNum());
                        // Иначе устанавливаем значение ширины ячейки как ширина заголовка * 256. Проверка на не null
                    } else if (sheet.getRow(i).getCell(k).getStringCellValue() != null && sheet.getRow(i).getCell(k).getStringCellValue().length() <= sheet.getRow(2).getCell(k).getStringCellValue().length()){
                        sheet.setColumnWidth(k, sheet.getRow(2).getCell(k).getStringCellValue().length() * 256);
                    }
                }
            }
        }
        //Запись выходного файла. Название берем из заголовка
        FileOutputStream fileOut = new FileOutputStream(parsingHtml.getTitle() + ".xls");
        wb.write(fileOut);
        fileOut.close();
    }
}