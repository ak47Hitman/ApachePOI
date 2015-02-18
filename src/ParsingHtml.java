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

    public int getNumberOfTDForMergeSells() {
        int k =0;
        Document doc = Jsoup.parse(inputText);
        Elements d = doc.getElementsByTag("table");
        for (Element a : d) {
            Elements trs = a.getElementsByTag("tr");
            for (Element tr : trs) {
                System.out.println("TR: " + tr.text());
                int i = 0;
                for (Element td : tr.getAllElements()) {
                    if (i != 0) {
                        k++;
                        System.out.println("TD: " + td.text());
                    }
                    i++;
                }
                break;
            }
        }
        return k;
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
//        Element table = doc.getElementsByTag("table").first();
//        Elements trs = table.getElementsByTag("tr");
//        for (Element tr : trs) {
//            System.out.println("TR: " + tr.text());
//            for (Element td : tr.getAllElements()) {
//                System.out.println("TD: " + td.text());
//            }
//        }
        System.out.println();
    }
//    public void takeTable() {
////        System.out.println(inputText);
//        System.out.println("*** JSOUP ***");
//        Document doc = Jsoup.parse(inputText);
//        System.out.println("H3: " + doc.getElementsByTag("h3").text());
////        System.out.println("Title: " + doc.getElementsByTag("title").text());
////        System.out.println("H1: " + doc.getElementsByTag("h1").text());
//        Element table = doc.getElementsByTag("table").first();
//        Elements trs = table.getElementsByTag("tr");
//        for (Element tr : trs) {
//            System.out.println("TR: " + tr.text());
//            for (Element td : tr.getAllElements()) {
//                System.out.println("TD: " + td.text());
//            }
//        }
//        System.out.println();
//    }
}
