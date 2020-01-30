package cn.hillwind.tools.excel;

public class Templates {

    static final String ContentTypeBegin = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n" +
            "    <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n" +
            "    <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n" +
            "    <Override PartName=\"/xl/workbook.xml\"\n" +
            "              ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n" +
            "    <Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\n" +
            "    <Override PartName=\"/xl/styles.xml\"\n" +
            "              ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n" +
            "    <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\n" +
            "    <Override PartName=\"/docProps/app.xml\"\n" +
            "              ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>\n";
    static final String ContentTypeItem = "    <Override PartName=\"/xl/worksheets/sheet%d.xml\"\n ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n";
    static final String ContentTypeEnd = "</Types>\n";

    // xl/workbook.xml
    static final String WorkbookBegin = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<workbook xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"\n" +
            "          xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n" +
            "    <fileVersion appName=\"xl\" lastEdited=\"5\" lowestEdited=\"5\" rupBuild=\"9303\"/>\n" +
            "    <workbookPr filterPrivacy=\"1\" defaultThemeVersion=\"124226\"/>\n" +
            "    <bookViews>\n" +
            "        <workbookView xWindow=\"240\" yWindow=\"120\" windowWidth=\"16155\" windowHeight=\"8505\"/>\n" +
            "    </bookViews>\n" +
            "    <sheets>";
    static final String WorkbookSheetItem = "<sheet name=\"%s\" sheetId=\"%d\" r:id=\"rId%d\"/>";
    static final String WorkbookEnd = "</sheets>\n" +
            "    <calcPr calcId=\"145621\"/>\n" +
            "</workbook>";

    // xl/_rels/workbook.xml.rels
    static final String XlRelBegin = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n";
    static final String XlRelTop1 = "<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>\n";
    static final String XlRelTop2 = "<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\n";
    static final String XlRelItem = "<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet%d.xml\"/>\n";
    static final String XlRelEnd = "</Relationships>";

    static final String SheetContentBegin = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<worksheet xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"\n" +
            "           xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"\n" +
            "           xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"\n" +
            "           mc:Ignorable=\"x14ac\">\n" +
            "    <dimension ref=\"A1\"/>\n" +
            "    <sheetViews>\n" +
            "        <sheetView tabSelected=\"1\" workbookViewId=\"0\">\n" +
            "            <selection activeCell=\"K18\" sqref=\"K18\"/>\n" +
            "        </sheetView>\n" +
            "    </sheetViews>\n" +
            "    <sheetFormatPr defaultRowHeight=\"13.5\" x14ac:dyDescent=\"0.15\"/>\n";

    static final String SheetContentEnd = "\n" +
            "    <phoneticPr fontId=\"1\" type=\"noConversion\"/>\n" +
            "    <pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>\n" +
            "</worksheet>";

}
