package cn.hillwind.tools.excel;

import java.io.*;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class Workbook implements AutoCloseable {

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
    }

    public Sheet createSheet(String name, int... columnWidths) throws IOException {
        int index = sheetList.size() + 1;
        String fileName = String.format("xl/worksheets/sheet%d.xml", index);
        File sheetFile = new File(rootFolder, fileName);
        sheetFile.getParentFile().mkdirs();
        Sheet sheet = new Sheet(index, name, sheetFile, columnWidths);
        fileNames.add(fileName);
        sheetList.add(sheet);
        return sheet;
    }

    @Override
    public void close() throws IOException {
        writeFiles();
        zip();
        clean();
    }

    private void writeFiles() throws IOException {
        copyResource("_rels/.rels");
        copyResource("docProps/app.xml");
        copyResource("docProps/core.xml");
        copyResource("xl/styles.xml");
        copyResource("xl/theme/theme1.xml");
//        copyResource("[Content_Types].xml");
        addFile("[Content_Types].xml", assetContentType());
//        copyResource("xl/workbook.xml");
        addFile("xl/workbook.xml", assetWorkbook());

//        copyResource("xl/_rels/workbook.xml.rels");
        addFile("xl/_rels/workbook.xml.rels", assetXlRel());

//        copyResource("xl/worksheets/sheet1.xml");
/*
        for(Sheet sheet: sheetList){
            addFile("xl/worksheets/sheet" + sheet.getIndex() + ".xml", assetSheet());
        }
*/
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

    private void addFile(String path, String content) throws IOException {
        fileNames.add(path);
        File targetFile = new File(rootFolder, path);
        targetFile.getParentFile().mkdirs();
        Files.write(targetFile.toPath(), content.getBytes());
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

    private void write(String path, String content) throws IOException {
        File f = new File(rootFolder, path);
        f.getParentFile().mkdirs();
        FileWriter fw = new FileWriter(f);
        fw.write(content);
        fw.close();
    }
}
