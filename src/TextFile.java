import java.io.*;
import java.util.ArrayList;

/**
 * Created by xhplogvievg_ on 04.02.2015.
 */

public class TextFile {
    // Считываем файл и выдаем строку на результат
    public static String read(String fileName) {
        StringBuilder sb = new StringBuilder();
        try {
            BufferedReader in = new BufferedReader(new FileReader( new File(fileName).getAbsoluteFile()));
            try {
                String s;
                while ((s = in.readLine()) != null) {
                    sb.append(s);
                    sb.append("\n");
                }
            } finally {
                in.close();
            }
        } catch(IOException e) {
            System.out.println("File does not exist: " + e);
            throw new RuntimeException(e);
        }
        return sb.toString();
    }
}