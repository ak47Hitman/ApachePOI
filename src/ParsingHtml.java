import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

/**
 * Created by logvinov on 18.02.2015.
 */
public class ParsingHtml {
    private String inputText;
    // Конструктор, в который мы передаем текст из HTML-файла
    public ParsingHtml(String inputText) {
        this.inputText = inputText;
    }

    // Получаем заголовок из нашего Html-файла
    // Заголовок заключен в тэг <H3>
    public String getHeader() {
        System.out.println("*** JSOUP ***");
        Document doc = Jsoup.parse(inputText);
        System.out.println("H3: " + doc.getElementsByTag("h3").text());
        return doc.getElementsByTag("h3").text();
    }
    // Получаем название файла
    // Название заклчено в теги <title></title>
    public String getTitle() {
        Document doc = Jsoup.parse(inputText);
        System.out.println("Title: " + doc.getElementsByTag("title").text());
        return doc.getElementsByTag("title").text();
    }
    // Получаем число объединенных ячеек
    // Оно равно максимальному числу ячеек среди всех строк в таблице
    // Пригодится при подсчете объединенных ячеек для задания красивого заголовка
    public int getNumberOfTDForMergeCells() {
        int maxNumberOfCellsInRow = 0;
        int numberOfCellsInRow = 0;
        Document doc = Jsoup.parse(inputText);
        // Просматриваем все таблицы в ячейке
        Elements tables = doc.getElementsByTag("table");
        for (Element table : tables) {
            // Для каждой выбранной таблицы выбираем строки <tr> и в них начинаем пробегать по каждой строке
            Elements trs = table.getElementsByTag("tr");
            for (Element tr : trs) {
                // Для каждой строки выбираем ячейку в строке и накручиваем счетчик
                for (Element tds : tr.getElementsByTag("td")) {
                    numberOfCellsInRow++;
                }
                // Если максимальное значение ячеек в таблице меньше нашего накрученного счетчика, то обновляем его
                if( maxNumberOfCellsInRow < numberOfCellsInRow) {
                    maxNumberOfCellsInRow = numberOfCellsInRow;
                }
                // Не забываем обнулить счетчик ячеек
                numberOfCellsInRow = 0;
            }
        }
        System.out.println(maxNumberOfCellsInRow);
        // Возвращаем максимальное значение ячеек
        return maxNumberOfCellsInRow;
    }
}
