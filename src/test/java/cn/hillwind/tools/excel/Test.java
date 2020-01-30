package cn.hillwind.tools.excel;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Test {
    public static void main(String[] args) throws IOException {
        benchmark();
        test();
    }

    public static void test() throws IOException {
        String filePath = "/tmp/test.xlsx";
        File excelFile = new File(filePath);
        try (Workbook workbook = new Workbook(excelFile)) {
            for (int i = 1; i < 6; i++) {
                try (Sheet sheet = workbook.createSheet("TestSheet" + i, 30, 20, 10)) { // width: 30,20,10 chars
                    sheet.blankRow();
                    sheet.row("<name>Outersky</name>", "&Amy", "'0123'", "\"123\"");
                    sheet.blankRows(5);
                    String[] line = new String[]{"1", "2", "3"};
                    sheet.row(line);
                }
            }
        }
    }

    public static void benchmark() throws IOException {
        long time = System.currentTimeMillis();
        String filePath = "/tmp/benchmark.xlsx";
        File excelFile = new File(filePath);

        try (Workbook workbook = new Workbook(excelFile)) {
            for (int i = 1; i < 11; i++) {
                String[] row = mockRow("SheetData" + i + ":");
                try (Sheet sheet = workbook.createSheet("Sheet" + i, 30, 20, 10)) { // width: 30,20,10 chars
                    for (int j = 0; j < 20000; j++) {
                        sheet.blankRow();
                        sheet.row(row);
                    }
                }
            }
        }
        System.out.println("time used(ms): " + (System.currentTimeMillis() - time));
    }

    private static String[] mockRow(String prefix) {
        List<String> data = new ArrayList<String>();
        for (int i = 0; i < 20; i++) {
            data.add(prefix + (i + 1));
        }
        String[] row = new String[0];
        row = data.toArray(row);
        return row;
    }
}
