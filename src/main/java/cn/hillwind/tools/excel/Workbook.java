package cn.hillwind.tools.excel;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class Workbook implements AutoCloseable {

    private static final String ResourcePathPrefix = "sample/";
    private List<Sheet> sheetList = new ArrayList<>();

    private FileOutputStream fos;
    private ZipOutputStream zout;

    //shared within the Workbook
    protected StringBuilder stringBuilder = new StringBuilder();

    /**
     * Create a Workbook instance
     *
     * @param excelFile excel file
     * @throws IOException exception
     */
    public Workbook(File excelFile) throws IOException {
        fos = new FileOutputStream(excelFile);
        zout = new ZipOutputStream(new BufferedOutputStream(fos, 100 * 1024));
        addHeadZipEntries();
    }

    /**
     * create a sheet
     *
     * @param name         sheet name
     * @param columnWidths sheet column widths
     * @return Sheet
     * @throws IOException exception
     */
    public Sheet createSheet(String name, int... columnWidths) throws IOException {
        int index = sheetList.size() + 1;
        String path = String.format("xl/worksheets/sheet%d.xml", index);
        zout.putNextEntry(new ZipEntry(path));
        Sheet sheet = new Sheet(this, index, name, zout, columnWidths);
        sheetList.add(sheet);
        return sheet;
    }

    @Override
    public void close() throws IOException {
        addZipEntry("[Content_Types].xml", assetContentType());
        addZipEntry("xl/workbook.xml", assetWorkbook());
        addZipEntry("xl/_rels/workbook.xml.rels", assetXlRel());
        zout.close();
        fos.close();

        clear();
    }

    private void clear() {
        stringBuilder = null;
        zout = null;
        fos = null;
        sheetList.clear();
        sheetList = null;
    }

    private void addHeadZipEntries() throws IOException {
        addResourceZipEntry("_rels/.rels");
        addResourceZipEntry("docProps/app.xml");
        addResourceZipEntry("docProps/core.xml");
        addResourceZipEntry("xl/styles.xml");
        addResourceZipEntry("xl/theme/theme1.xml");
    }

    private void addZipEntry(String path, String content) throws IOException {
        ZipEntry entry = new ZipEntry(path);
        zout.putNextEntry(entry);
        zout.write(content.getBytes());
        zout.closeEntry();
    }

    private void addResourceZipEntry(String path) throws IOException {
        ZipEntry entry = new ZipEntry(path);
        zout.putNextEntry(entry);
        InputStream stream = Workbook.class.getResourceAsStream(ResourcePathPrefix + path);
        byte[] buffer = new byte[1024];
        int len;
        while ((len = stream.read(buffer)) != -1) {
            zout.write(buffer, 0, len);
        }
        zout.closeEntry();
    }

    // xl/workbook.xml
    private String assetContentType() {
        StringBuilder sb = new StringBuilder(Templates.ContentTypeBegin);
        for (Sheet sheet : sheetList) {
            sb.append(String.format(Templates.ContentTypeItem, sheet.getIndex()));
        }
        sb.append(Templates.ContentTypeEnd);
        return sb.toString();
    }

    // xl/workbook.xml
    private String assetWorkbook() {
        StringBuilder sb = new StringBuilder(Templates.WorkbookBegin);
        for (Sheet sheet : sheetList) {
            sb.append(String.format(Templates.WorkbookSheetItem, sheet.getName(), sheet.getIndex(), sheet.getIndex()));
        }
        sb.append(Templates.WorkbookEnd);
        return sb.toString();
    }

    // xl/_rels/workbook.xml.rels
    private String assetXlRel() {
        StringBuilder sb = new StringBuilder(Templates.XlRelBegin);
        sb.append(String.format(Templates.XlRelTop1, sheetList.size() + 2));
        sb.append(String.format(Templates.XlRelTop2, sheetList.size() + 1));
        for (int i = sheetList.size(); i > 0; i--) {
            sb.append(String.format(Templates.XlRelItem, i, i));
        }
        sb.append(Templates.XlRelEnd);
        return sb.toString();
    }

}
