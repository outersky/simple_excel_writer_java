package cn.hillwind.tools.excel;

import java.io.*;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class Workbook {

    private static final String ResourcePathPrefix = "sample/";
    private List<Sheet> sheetList = new ArrayList<>();
    private File rootFolder;
    private File excelFile;
    private Set<String> fileNames = new HashSet<>();

    Workbook(File excelFile) throws IOException {
        this.excelFile = excelFile;
        rootFolder = File.createTempFile("temp_excel_dir", "");
        rootFolder.delete();
        rootFolder.mkdir();
        System.out.println("temp dir created: " + rootFolder.getAbsolutePath());
    }

    public static void test(File excelFile) throws IOException {
        Workbook workbook = new Workbook(excelFile);
        workbook.close();
    }

    public Sheet createSheet(String name) throws IOException {
        int index = sheetList.size() + 1;
        String fileName = String.format("xl/worksheets/sheet%d.xml", index);
        File sheetFile = new File(rootFolder, fileName);
        sheetFile.getParentFile().mkdirs();
        Sheet sheet = new Sheet(index, name, sheetFile);
        sheetList.add(sheet);
        return sheet;
    }

    public void close() throws IOException {
        writeFiles();
        zip();
        clean();
    }

    private void writeFiles() throws IOException {
        copyResource("[Content_Types].xml");
        copyResource("_rels/.rels");
        copyResource("docProps/app.xml");
        copyResource("docProps/core.xml");
        copyResource("xl/styles.xml");
        copyResource("xl/workbook.xml");
        copyResource("xl/_rels/workbook.xml.rels");
        copyResource("xl/theme/theme1.xml");
        copyResource("xl/worksheets/sheet1.xml");
//        copyResource("xl/worksheets/sheet2.xml");
//        copyResource("xl/worksheets/sheet3.xml");
//        copyResource("xl/worksheets/_rels/sheet1.xml.rels");
//        copyResource("xl/worksheets/_rels/sheet2.xml.rels");
//        copyResource("xl/worksheets/_rels/sheet3.xml.rels");
    }

    private void zip() throws IOException {
        FileOutputStream fos = new FileOutputStream(excelFile);
        ZipOutputStream zout = new ZipOutputStream(fos);

        for (String path : fileNames) {
            addZipEntry(zout, path);
        }
        zout.closeEntry();
        zout.close();
        fos.close();
    }

    private void clean() throws IOException {
        for (String path : fileNames) {
            new File(rootFolder, path).delete();
        }
        rootFolder.delete();
    }

    private void addZipEntry(ZipOutputStream zipOutputStream, String path) throws IOException {
        File f = new File(rootFolder, path);
        FileInputStream fis = new FileInputStream(f);
        ZipEntry entry = new ZipEntry(path);
        zipOutputStream.putNextEntry(entry);

        byte[] buffer = new byte[1024];
        int len = -1;
        while ((len = fis.read(buffer)) != -1) {
            zipOutputStream.write(buffer, 0, len);
        }
    }

    private void copyResource(String path) throws IOException {
        fileNames.add(path);
        InputStream stream = Workbook.class.getResourceAsStream(ResourcePathPrefix + path);
        File targetFile = new File(rootFolder, path);
        targetFile.getParentFile().mkdirs();
        Files.copy(stream, targetFile.toPath());
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
    private String assetWorkbookRels() {
        StringBuilder sb = new StringBuilder(Templates.XlRelBegin);
        for (Sheet sheet : sheetList) {
            sb.append(String.format(Templates.XlRelItem, sheet.getIndex(), sheet.getIndex()));
        }
        sb.append(Templates.XlRelEnd);
        return sb.toString();
    }

    private void write(String path, String content) throws IOException {
        File f = new File(rootFolder, path);
        f.getParentFile().mkdirs();
        FileWriter fw = new FileWriter(f);
        fw.write(content);
        fw.close();
    }
}
