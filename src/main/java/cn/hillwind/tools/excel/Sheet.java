package cn.hillwind.tools.excel;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

public class Sheet {

    static final String SheetBegin = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"\n" +
            "        xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">";
    static final String SheetEnd = "</worksheet>";
    static final String SheetDataBegin = "<sheetData>";
    static final String SheetDataEnd = "</sheetData>";
    private final String name;
    private final int index;
    private final File file;
    private transient FileWriter writer;

    public Sheet(int index, String name, File file) throws IOException {
        this.index = index;
        this.name = name;
        this.file = file;
        writer = new FileWriter(file);

        writer.write(SheetBegin);
//        writer.write(SheetDataBegin);
    }

    public void close() throws IOException {
//        writer.write(SheetDataEnd);
        writer.write(SheetEnd);
        writer.close();
    }

    public String getName() {
        return name;
    }

    public int getIndex() {
        return index;
    }

    public File getFile() {
        return file;
    }

}
