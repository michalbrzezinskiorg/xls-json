package xls.e2j;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;
import org.xml.sax.SAXException;

import java.io.*;

public class ExcelToJsonTest {

    public static final String SRC_MAIN_RESOURCES = "./src/resources/";
    ExcelToJson de = new ExcelToJson();

    @Test
    public void shouldFail() {
        File f = new File(SRC_MAIN_RESOURCES+"test.numbers");
        Assertions.assertThrows(IllegalArgumentException.class , () -> de.getJsonString(new BufferedInputStream(new FileInputStream(f))));
    }

    @Test
    public void xls() throws IOException, SAXException, OpenXML4JException {
        File f = new File(SRC_MAIN_RESOURCES+"testXLS.xls");
        BufferedInputStream inputStream = new BufferedInputStream(new FileInputStream(f));
        String res = de.getJsonString(inputStream);
        System.out.println("xls: \n"+res);
    }

    @Test
    public void xlsx() throws IOException, SAXException, OpenXML4JException {
        File f = new File(SRC_MAIN_RESOURCES+"testXLSX.xlsx");
        String res = de.getJsonString(new BufferedInputStream(new FileInputStream(f)));
        System.out.println("xlsx \n: "+res);
    }

}