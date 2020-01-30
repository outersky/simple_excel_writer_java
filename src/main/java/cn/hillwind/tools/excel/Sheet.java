package cn.hillwind.tools.excel;

import java.io.BufferedOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.zip.ZipOutputStream;

public class Sheet implements AutoCloseable {

    static final String SheetDataBegin = "<sheetData>\n";
    static final String SheetDataEnd = "</sheetData>\n";
    private final String name;
    private final int index;

    ZipOutputStream zout;
    Writer writer;
    private int rowIndex = 0;

    /**
     * 创建Sheet
     *
     * @param index        索引
     * @param name         名称
     * @param zout         输出压缩流
     * @param columnWidths 列宽（以字符为单位）
     * @throws IOException 错误
     */
    protected Sheet(int index, String name, ZipOutputStream zout, int... columnWidths) throws IOException {
        this.index = index;
        this.name = name;
        this.zout = zout;
        writer = new OutputStreamWriter(new BufferedOutputStream(zout));

        writer.write(Templates.SheetContentBegin);
        if (columnWidths != null && columnWidths.length > 0) {
            writer.write("<cols>");
            for (int i = 0; i < columnWidths.length; i++) {
                writer.write(String.format("<col min=\"%d\" max=\"%d\" width=\"%d.7\" customWidth=\"1\"/>\n", i + 1, i + 1, columnWidths[i]));
            }
            writer.write("</cols>");
        }
        writer.write(SheetDataBegin);
    }

    public void row(String... args) throws IOException {
        rowIndex++;
        writer.write(String.format("<row r=\"%d\" spans=\"1:%d\">\n", rowIndex, args.length));
        for (int j = 0; j < args.length; j++) {
            writer.write(String.format("<c r=\"%s%d\" t=\"str\"><v>%s</v></c>", colStr(j + 1), rowIndex, args[j]));
        }
        writer.write("</row>\n");
    }

    public void blankRow() {
        blankRows(1);
    }

    public void blankRows(int rows) {
        if (rows <= 0) return;
        rowIndex += rows;
    }

    /**
     * 将数值的列下标缓存字符串，如 1 -> A
     *
     * @param colIndex 1开始的下标
     * @return 字符串列名， 如 AA ~ FZ
     */
    private String colStr(int colIndex) {
        colIndex--; // 换成0计数
        StringBuilder sb = new StringBuilder();
        char b;
        while (colIndex >= 0) {
            b = (char) ('A' + (colIndex % 26));
            sb.append(b);
            colIndex = colIndex / 26 - 1;
        }
        return sb.reverse().toString();
    }

    @Override
    public void close() throws IOException {
        writer.write(SheetDataEnd);
        writer.write(Templates.SheetContentEnd);
        writer.flush();
        zout.closeEntry();
    }

    public String getName() {
        return name;
    }

    public int getIndex() {
        return index;
    }

}
