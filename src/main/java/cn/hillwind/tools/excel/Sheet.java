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

    private ZipOutputStream zout;
    private Writer writer;
    private int rowIndex = 0;
    private Workbook workbook;

    /**
     * 创建Sheet
     *
     * @param index        索引
     * @param name         名称
     * @param zout         输出压缩流
     * @param columnWidths 列宽（以字符为单位）
     * @throws IOException 错误
     */
    protected Sheet(Workbook workbook, int index, String name, ZipOutputStream zout, int... columnWidths) throws IOException {
        this.workbook = workbook;
        this.index = index;
        this.name = name;
        this.zout = zout;
        writer = new OutputStreamWriter(new BufferedOutputStream(zout, 100 * 1024));

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

    /**
     * append a row to sheet
     *
     * @param values cell values
     * @throws IOException exception
     */
    public void row(String... values) throws IOException {
        rowIndex++;
        // reuse StringBuilder
        StringBuilder stringBuilder = workbook.stringBuilder;
        stringBuilder.setLength(0);
        stringBuilder.append("<row r=\"").append(rowIndex).append("\" spans=\"1:").append(values.length).append("\">\n");
        for (int j = 0; j < values.length; j++) {
            stringBuilder.append("<c r=\"").append(colStr(j + 1)).append(rowIndex).append("\" t=\"str\"><v>").append(xmlEscape(values[j])).append("</v></c>");
        }
        stringBuilder.append("</row>\n");
        writer.write(stringBuilder.toString());
        stringBuilder.setLength(0);
    }

    private String xmlEscape(String value) {
        if (value == null) return "";
        if (value.length() == 0) return value;
        return value.replaceAll("&", "&amp;")
                .replaceAll("<", "&lt;")
                .replaceAll(">", "&gt;")
                .replaceAll("'", "&apos;")
                .replaceAll("\"", "&quot;");
    }

    /**
     * append a blank row
     */
    public void blankRow() {
        blankRows(1);
    }

    /**
     * append several blank rows
     *
     * @param count count of blank rows
     */
    public void blankRows(int count) {
        if (count <= 0) return;
        rowIndex += count;
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
        clear();
    }

    private void clear() {
        writer = null;
        zout = null;
        workbook = null;
    }

    public String getName() {
        return name;
    }

    public int getIndex() {
        return index;
    }

}
