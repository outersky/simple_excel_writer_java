package cn.hillwind.tools.excel;

import java.io.File;
import java.io.IOException;

public class Test {
    public static void main(String[] args) throws IOException {
        String filePath = "/tmp/test.xlsx";
        File excelFile = new File(filePath);
        try (Workbook workbook = new Workbook(excelFile)) {
            for (int i = 1; i < 6; i++) {
                try (Sheet sheet = workbook.createSheet("TestSheet" + i, 30, 20, 10)) { // width: 30,20,10 chars
                    sheet.blankRow();
                    sheet.row("A", "B", "C");
                    sheet.blankRows(5);
                    String[] line = new String[]{"1", "2", "3"};
                    sheet.row(line);
                }
            }
        }
    }
}
