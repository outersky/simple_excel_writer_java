package cn.hillwind.tools.excel;

import java.io.File;
import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {
        long time = System.currentTimeMillis();
        try {
            Workbook.test(new File("/tmp/abc.xlsx"));
        } finally {
            System.out.println("time used (ms): " + (System.currentTimeMillis() - time));
        }
    }
}
