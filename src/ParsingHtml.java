import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

/**
 * Created by logvinov on 18.02.2015.
 */
public class ParsingHtml {
    private String inputText;
    public ParsingHtml(String inputText) {
        this.inputText = inputText;
    }
    public String getHeader() {
        System.out.println("*** JSOUP ***");
        Document doc = Jsoup.parse(inputText);
        System.out.println("H3: " + doc.getElementsByTag("h3").text());
        return doc.getElementsByTag("h3").text();
    }

    public String getTitle() {
        Document doc = Jsoup.parse(inputText);
        System.out.println("Title: " + doc.getElementsByTag("title").text());
        return doc.getElementsByTag("title").text();
    }

    public int getNumberOfTDForMergeCells() {
        int maxNumberOfCellsInRow = 0;
        int numberOfCellsInRow = 0;
        Document doc = Jsoup.parse(inputText);
        Elements tables = doc.getElementsByTag("table");
        for (Element table : tables) {
            Elements trs = table.getElementsByTag("tr");
            for (Element tr : trs) {
                for (Element tds : tr.getElementsByTag("td")) {
                            numberOfCellsInRow++;
                }
                if( maxNumberOfCellsInRow < numberOfCellsInRow) {
                    maxNumberOfCellsInRow = numberOfCellsInRow;
                }
                numberOfCellsInRow = 0;
            }
        }
        System.out.println(maxNumberOfCellsInRow);
        return maxNumberOfCellsInRow;
    }

    public void getTable() {
        Document doc = Jsoup.parse(inputText);
        Elements d = doc.getElementsByTag("table");
        for (Element a : d) {
            Elements trs = a.getElementsByTag("tr");
            for (Element tr : trs) {
                System.out.println("TR: " + tr.text());
                int i = 0;
                for (Element td : tr.getAllElements()) {
                    if (i != 0) {
                        System.out.println("TD: " + td.text());
                    }
                    i++;
                }
            }
        }
        System.out.println();
    }
}
