package xls.j2e;

import org.junit.jupiter.api.Test;

import java.io.BufferedOutputStream;
import java.io.IOException;

class JsonToExcelTest {

    JsonToExcel jsonToExcel = new JsonToExcel();
    String json = "{\"Arkusz 1\":{\"0\":{\"0\":\"Tabela 1\"},\"1\":{\"0\":\"zxv1\",\"1\":\"Vas2\",\"2\":\"Xcb3\",\"3\":\"Erg4\",\"4\":\"Asc5\",\"5\":\"Few6\",\"6\":\"Qwef7\"},\"2\":{\"0\":\"qwd\",\"1\":\"qwdw\",\"2\":\"dw\",\"3\":\"asda\",\"4\":\"va\",\"5\":\"234\",\"6\":\"234\"},\"3\":{\"0\":\"e\",\"1\":\"wer\",\"2\":\"w\",\"3\":\"er\",\"4\":\"2\",\"5\":\"234\",\"6\":\"34\"},\"4\":{\"0\":\"werwer\",\"1\":\"eerw\",\"2\":\"werwe\",\"3\":\"rewerw\",\"4\":\"erwe\",\"5\":\"456\",\"6\":\"564\"},\"5\":{\"0\":\"tyjyth\",\"1\":\"jty\",\"2\":\"dfgdf\",\"3\":\"tyj\",\"4\":\"rert\",\"5\":\"5545\",\"6\":\"456\"},\"6\":{\"0\":\"htjyj\",\"1\":\"yjty\",\"2\":\"gdfgdjrty\",\"3\":\"eerted\",\"4\":\"erter\",\"5\":\"6456\",\"6\":\"565\"},\"7\":{\"5\":\"12925\",\"6\":\"1853\"}}}";
    public static final String SRC_MAIN_RESOURCES = "./src/resources/";

    @Test
    public void getFile() throws IOException {
        BufferedOutputStream test = jsonToExcel.getXls(json, SRC_MAIN_RESOURCES+"testResult.xls");
        test.flush();
    }

}