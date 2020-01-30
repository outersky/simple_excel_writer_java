package cn.hillwind.tools.excel;

public class Templates {
    // xl/workbook.xml
    static final String WorkbookBegin = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n" +
            "    <workbookPr date1904=\"false\"/>\n" +
            "    <sheets>\n";
    static final String WorkbookSheetItem = "<sheet name=\"%s\" sheetId=\"%d\" r:id=\"rId%d\"/>";
    static final String WorkbookEnd = "</sheets>\n" +
            "</workbook>";

    // xl/_rels/workbook.xml.rels
    static final String XlRelBegin = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>\n" +
            "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>";
    static final String XlRelItem = "<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet%d.xml\"/>";
    static final String XlRelEnd = "</Relationships>";

}
