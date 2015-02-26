import java.io.*;

public class Main {

    public static void main(String[] args) throws IOException {
        // Задаем HTML - файл для парсинга
        String htmlFileName = "C:\\Users\\logvinov\\Desktop\\MyWorkSpace\\ApachePOI\\1.xls";
        // Создаем экземпляр класса для чтения текста из файла
        TextFile htmlFile = new TextFile();
        // Создаем экземпляр класса для передачи туда полученного текста и выгрузки его в Excel - файл
        CreateExcel createExcel = new CreateExcel();
        // Считываем текст из файла и присваиваем его в строке
        String htmlTextFromFile = htmlFile.read(htmlFileName);
        // Передаем полученный текст из файла
        createExcel.createExcelDoc(htmlTextFromFile);
        System.out.println("All done");
    }
}
