package com.ow.util;

import java.io.*;
import java.util.Formatter;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.examples.html.HSSFHtmlHelper;
import org.apache.poi.ss.examples.html.HtmlHelper;
import org.apache.poi.ss.examples.html.XSSFHtmlHelper;
import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.format.CellFormatResult;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelToHtml {

    private final Workbook wb;
    private final Appendable output;
    private boolean completeHTML;
    private Formatter out;
    private boolean gotBounds;
    private int firstColumn;
    private int endColumn;
    private HtmlHelper helper;

    private static final String DEFAULTS_CLASS = "excelDefaults";
    private static final String COL_HEAD_CLASS = "colHeader";
    private static final String ROW_HEAD_CLASS = "rowHeader";

    private static final Map<HorizontalAlignment, String> HALIGN = mapFor(
            HorizontalAlignment.LEFT, "left",
            HorizontalAlignment.CENTER, "center",
            HorizontalAlignment.RIGHT, "right",
            HorizontalAlignment.FILL, "left",
            HorizontalAlignment.JUSTIFY, "left",
            HorizontalAlignment.CENTER_SELECTION, "center");

    private static final Map<VerticalAlignment, String> VALIGN = mapFor(
            VerticalAlignment.BOTTOM, "bottom",
            VerticalAlignment.CENTER, "middle",
            VerticalAlignment.TOP, "top");

    private static final Map<BorderStyle, String> BORDER = mapFor(
            BorderStyle.DASH_DOT, "dashed 1pt",
            BorderStyle.DASH_DOT_DOT, "dashed 1pt",
            BorderStyle.DASHED, "dashed 1pt",
            BorderStyle.DOTTED, "dotted 1pt",
            BorderStyle.DOUBLE, "double 3pt",
            BorderStyle.HAIR, "solid 1px",
            BorderStyle.MEDIUM, "solid 2pt",
            BorderStyle.MEDIUM_DASH_DOT, "dashed 2pt",
            BorderStyle.MEDIUM_DASH_DOT_DOT, "dashed 2pt",
            BorderStyle.MEDIUM_DASHED, "dashed 2pt",
            BorderStyle.NONE, "none",
            BorderStyle.SLANTED_DASH_DOT, "dashed 2pt",
            BorderStyle.THICK, "solid 3pt",
            BorderStyle.THIN, "dashed 1pt");

    private static final int IDX_TABLE_WIDTH = -2;
    private static final int IDX_HEADER_COL_WIDTH = -1;

    /**
     * Run this class as a program
     *
     * @param args The command line arguments.
     *
     * @throws Exception Exception we don't recover from.
     */
    public static void main(String[] args) throws Exception {
        if(args.length < 2){
            System.err.println("usage: ToHtml inputWorkbook outputHtmlFile");
            return;
        }

        try (
                FileWriter fw = new FileWriter(args[1]);
                PrintWriter pw = new PrintWriter(fw)
        ) {
            ExcelToHtml toHtml = create(args[0], pw);
            toHtml.setCompleteHTML(true);
            toHtml.printPage();
        }
    }

    @SuppressWarnings({"unchecked"})
    private static <K, V> Map<K, V> mapFor(Object... mapping) {
        Map<K, V> map = new HashMap<>();
        for (int i = 0; i < mapping.length; i += 2) {
            map.put((K) mapping[i], (V) mapping[i + 1]);
        }
        return map;
    }

    public static ExcelToHtml create(Workbook wb, Appendable output) {
        return new ExcelToHtml(wb, output);
    }

    public static ExcelToHtml create(String path, Appendable output)
            throws IOException {
        System.out.println("开始读取工作表，创建html文件...");
        return create(new FileInputStream(path), output);
    }

    public static ExcelToHtml create(InputStream in, Appendable output)
            throws IOException {
        Workbook wb = WorkbookFactory.create(in);
        return create(wb, output);
    }

    private ExcelToHtml(Workbook wb, Appendable output) {
        if (wb == null) {
            throw new NullPointerException("wb");
        }
        if (output == null) {
            throw new NullPointerException("output");
        }
        this.wb = wb;
        this.output = output;
        setupColorMap();
    }

    private void setupColorMap() {
        if (wb instanceof HSSFWorkbook) {
            helper = new HSSFHtmlHelper((HSSFWorkbook) wb);
        } else if (wb instanceof XSSFWorkbook) {
            helper = new XSSFHtmlHelper();
        } else {
            throw new IllegalArgumentException(
                    "unknown workbook type: " + wb.getClass().getSimpleName());
        }
    }

    public void setCompleteHTML(boolean completeHTML) {
        this.completeHTML = completeHTML;
    }

    public void printPage() throws IOException {
        System.out.println("开始输出内容到html文件中...");
        try {
            ensureOut();
            if (completeHTML) {
                out.format(
                        "<?xml version=\"1.0\" encoding=\"iso-8859-1\" ?>%n");
                out.format("<html>%n");
                out.format("<head>%n");
                out.format("</head>%n");
                out.format("<body>%n");
            }

            print();

            if (completeHTML) {
                out.format("</body>%n");
                out.format("</html>%n");
            }
            System.out.println("导入成功！");
        } finally {
            IOUtils.closeQuietly(out);
            if (output instanceof Closeable) {
                IOUtils.closeQuietly((Closeable) output);
            }
        }
    }

    public void print() {
        System.out.println("开始打印行格式...");
        printInlineStyle();
        System.out.println("开始打印工作表sheets...");
        printSheets();
    }

    private void printInlineStyle() {
        //out.format("<link href=\"excelStyle.css\" rel=\"stylesheet\" type=\"text/css\">%n");
        out.format("<style type=\"text/css\">%n");
        printStyles();
        out.format("</style>%n");
    }

    private void ensureOut() {
        if (out == null) {
            out = new Formatter(output);
        }
    }

    public void printStyles() {
        System.out.println("开始读取excelStyle.css文件...");
        InputStream InputStream = null;
        ensureOut();
        InputStream = ExcelToHtml.class.getClassLoader().getResourceAsStream("excelStyle.css");
        try (BufferedReader in = new BufferedReader(new InputStreamReader(InputStream))){
            String line;
            while ((line = in.readLine()) != null) {
                out.format("%s%n", line);
            }
        } catch (IOException e) {
            throw new IllegalStateException("Reading standard css", e);
        }

        // now add css for each used style
        Set<CellStyle> seen = new HashSet<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            Iterator<Row> rows = sheet.rowIterator();
            while (rows.hasNext()) {
                Row row = rows.next();
                for (Cell cell : row) {
                    CellStyle style = cell.getCellStyle();
                    if (!seen.contains(style)) {
                        printStyle(style);
                        seen.add(style);
                    }
                }
            }
        }
    }

    private void printStyle(CellStyle style) {
        out.format(".%s .%s {%n", DEFAULTS_CLASS, styleName(style));
        styleContents(style);
        out.format("}%n");
    }

    private void styleContents(CellStyle style) {
        styleOut("text-align", style.getAlignment(), HALIGN);
        styleOut("vertical-align", style.getVerticalAlignment(), VALIGN);
        fontStyle(style);
        borderStyles(style);
        helper.colorStyles(style, out);
    }

    private void borderStyles(CellStyle style) {
        styleOut("border-left", style.getBorderLeft(), BORDER);
        styleOut("border-right", style.getBorderRight(), BORDER);
        styleOut("border-top", style.getBorderTop(), BORDER);
        styleOut("border-bottom", style.getBorderBottom(), BORDER);
    }

    private void fontStyle(CellStyle style) {
        Font font = wb.getFontAt(style.getFontIndexAsInt());

        if (font.getBold()) {
            out.format("  font-weight: bold;%n");
        }
        if (font.getItalic()) {
            out.format("  font-style: italic;%n");
        }

        int fontheight = font.getFontHeightInPoints();
        if (fontheight == 9) {
            //fix for stupid ol Windows
            fontheight = 10;
        }
        out.format("  font-size: %dpt;%n", fontheight);

        // Font color is handled with the other colors
    }

    private String styleName(CellStyle style) {
        if (style == null) {
            style = wb.getCellStyleAt((short) 0);
        }
        StringBuilder sb = new StringBuilder();
        try (Formatter fmt = new Formatter(sb)) {
            fmt.format("style_%02x", style.getIndex());
            return fmt.toString();
        }
    }

    private <K> void styleOut(String attr, K key, Map<K, String> mapping) {
        String value = mapping.get(key);
        if (value != null) {
            out.format("  %s: %s;%n", attr, value);
        }
    }

    private static CellType ultimateCellType(Cell c) {
        CellType type = c.getCellType();
        if (type == CellType.FORMULA) {
            type = c.getCachedFormulaResultType();
        }
        return type;
    }

    private void printSheets() {
        ensureOut();
        Sheet sheet = wb.getSheetAt(0);
        printSheet(sheet);
    }

    public void printSheet(Sheet sheet) {
        ensureOut();
        Map<Integer, Integer> widths = computeWidths(sheet);
        int tableWidth = widths.get(IDX_TABLE_WIDTH);
        out.format("<table class=%s style=\"width:%dpx;\">%n", DEFAULTS_CLASS, tableWidth);
        printCols(widths);
        printSheetContent(sheet);
        out.format("</table>%n");
    }

    private Map<Integer, Integer> computeWidths(Sheet sheet) {
        Map<Integer, Integer> ret = new TreeMap<>();
        int tableWidth = 0;

        ensureColumnBounds(sheet);

        // compute width of the header column
        int lastRowNum = sheet.getLastRowNum();
        int headerCharCount = String.valueOf(lastRowNum).length();
        int headerColWidth = widthToPixels((headerCharCount + 1) * 256.0);
        ret.put(IDX_HEADER_COL_WIDTH, headerColWidth);
        tableWidth += headerColWidth;

        for (int i = firstColumn; i < endColumn; i++) {
            int colWidth = widthToPixels(sheet.getColumnWidth(i));
            ret.put(i, colWidth);
            tableWidth += colWidth;
        }

        ret.put(IDX_TABLE_WIDTH, tableWidth);
        return ret ;
    }

    protected int widthToPixels(final double widthUnits) {
        return Math.toIntExact(Math.round(widthUnits * 9 / 256));
    }

    private void printCols(Map<Integer, Integer> widths) {
        int headerColWidth = widths.get(IDX_HEADER_COL_WIDTH);
        out.format("<col style=\"width:%dpx\"/>%n", headerColWidth);
        for (int i = firstColumn; i < endColumn; i++) {
            int colWidth = widths.get(i);
            out.format("<col style=\"width:%dpx;\"/>%n", colWidth);
        }
    }

    private void ensureColumnBounds(Sheet sheet) {
        if (gotBounds) {
            return;
        }

        Iterator<Row> iter = sheet.rowIterator();
        firstColumn = (iter.hasNext() ? Integer.MAX_VALUE : 0);
        endColumn = 0;
        while (iter.hasNext()) {
            Row row = iter.next();
            short firstCell = row.getFirstCellNum();
            if (firstCell >= 0) {
                firstColumn = Math.min(firstColumn, firstCell);
                endColumn = Math.max(endColumn, row.getLastCellNum());
            }
        }
        gotBounds = true;
    }

    private void printColumnHeads() {
        out.format("<thead>%n");
        out.format("  <tr class=%s>%n", COL_HEAD_CLASS);
        out.format("    <th class=%s>&#x25CA;</th>%n", COL_HEAD_CLASS);
        //noinspection UnusedDeclaration
        StringBuilder colName = new StringBuilder();
        for (int i = firstColumn; i < endColumn; i++) {
            colName.setLength(0);
            int cnum = i;
            do {
                colName.insert(0, (char) ('A' + cnum % 26));
                cnum /= 26;
            } while (cnum > 0);
            out.format("    <th class=%s>%s</th>%n", COL_HEAD_CLASS, colName);
        }
        out.format("  </tr>%n");
        out.format("</thead>%n");
    }

    private void printSheetContent(Sheet sheet) {
        printColumnHeads();

        out.format("<tbody>%n");
        Iterator<Row> rows = sheet.rowIterator();
        while (rows.hasNext()) {
            Row row = rows.next();

            out.format("  <tr>%n");
            out.format("    <td class=%s>%d</td>%n", ROW_HEAD_CLASS,
                    row.getRowNum() + 1);
            for (int i = firstColumn; i < endColumn; i++) {
                String content = "&nbsp;";
                String attrs = "";
                CellStyle style = null;
                if (i >= row.getFirstCellNum() && i < row.getLastCellNum()) {
                    Cell cell = row.getCell(i);
                    if (cell != null) {
                        style = cell.getCellStyle();
                        attrs = tagStyle(cell, style);
                        //Set the value that is rendered for the cell
                        //also applies the format
                        CellFormat cf = CellFormat.getInstance(
                                style.getDataFormatString());
                        CellFormatResult result = cf.apply(cell);
                        content = result.text; //never null
                        if (content.isEmpty()) {
                            content = "&nbsp;";
                        }
                    }
                }
                out.format("    <td class=%s %s>%s</td>%n", styleName(style),
                        attrs, content);
            }
            out.format("  </tr>%n");
        }
        out.format("</tbody>%n");
    }

    private String tagStyle(Cell cell, CellStyle style) {
        if (style.getAlignment() == HorizontalAlignment.GENERAL) {
            switch (ultimateCellType(cell)) {
                case STRING:
                    return "style=\"text-align: left;\"";
                case BOOLEAN:
                case ERROR:
                    return "style=\"text-align: center;\"";
                case NUMERIC:
                default:
                    // "right" is the default
                    break;
            }
        }
        return "";
    }
}
