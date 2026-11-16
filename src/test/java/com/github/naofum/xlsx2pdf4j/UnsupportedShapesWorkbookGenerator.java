/*
 * Copyright 2026 Naofumi Fukue
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package com.github.naofum.xlsx2pdf4j;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.main.STShapeType;

import java.io.OutputStream;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Builds a catalog workbook containing every preset shape that
 * {@code ShapeRenderer} does not draw natively and therefore approximates
 * with its bounding rectangle.
 *
 * <p>Connector presets are intentionally excluded: they are rendered through
 * the connector path, not as {@code XSSFSimpleShape} presets. Custom geometry
 * is also outside this catalog.</p>
 */
public final class UnsupportedShapesWorkbookGenerator {

    public static final Path DEFAULT_OUTPUT =
            Path.of("testdata", "validation", "unsupported-shapes.xlsx");

    private static final int TILES_PER_ROW = 5;
    private static final int TILES_PER_PAGE = 25;
    private static final int TILE_ROWS = 4;
    private static final int FIRST_TILE_ROW = 3;

    /**
     * Presets handled by {@code ShapeRenderer#getShapeOutline}. Connector
     * constants are handled separately and are filtered out before this set
     * is applied.
     */
    private static final Set<String> SUPPORTED = Set.of(
            "LINE",
            "ARC",
            "BEVEL",
            "BLOCK_ARC",
            "BRACE_PAIR",
            "BRACKET_PAIR",
            "CAN",
            "CHART_PLUS",
            "CHART_STAR",
            "CHART_X",
            "CHEVRON",
            "CHORD",
            "CLOUD",
            "CORNER",
            "CORNER_TABS",
            "CUBE",
            "DECAGON",
            "DIAG_STRIPE",
            "DIAMOND",
            "DODECAGON",
            "DONUT",
            "DOUBLE_WAVE",
            "FRAME",
            "FUNNEL",
            "GEAR_6",
            "GEAR_9",
            "HALF_FRAME",
            "HEART",
            "HEPTAGON",
            "HEXAGON",
            "LEFT_BRACE",
            "LEFT_BRACKET",
            "LIGHTNING_BOLT",
            "LINE_INV",
            "MATH_DIVIDE",
            "MATH_EQUAL",
            "MATH_MINUS",
            "MATH_MULTIPLY",
            "MATH_NOT_EQUAL",
            "MATH_PLUS",
            "MOON",
            "NON_ISOSCELES_TRAPEZOID",
            "NO_SMOKING",
            "OCTAGON",
            "PARALLELOGRAM",
            "PENTAGON",
            "PIE",
            "PIE_WEDGE",
            "PLAQUE",
            "PLAQUE_TABS",
            "PLUS",
            "RIGHT_BRACE",
            "RIGHT_BRACKET",
            "RT_TRIANGLE",
            "SMILEY_FACE",
            "SQUARE_TABS",
            "SUN",
            "TEARDROP",
            "TRAPEZOID",
            "TRIANGLE",
            "WAVE",
            "ROUND_1_RECT",
            "ROUND_2_SAME_RECT",
            "ROUND_2_DIAG_RECT",
            "SNIP_1_RECT",
            "SNIP_2_SAME_RECT",
            "SNIP_2_DIAG_RECT",
            "SNIP_ROUND_RECT",
            "STAR_4",
            "STAR_6",
            "STAR_7",
            "STAR_8",
            "STAR_10",
            "STAR_12",
            "STAR_16",
            "STAR_24",
            "STAR_32",
            "IRREGULAR_SEAL_1",
            "CALLOUT_1",
            "CALLOUT_2",
            "CALLOUT_3",
            "ACCENT_CALLOUT_1",
            "ACCENT_CALLOUT_2",
            "ACCENT_CALLOUT_3",
            "ACCENT_BORDER_CALLOUT_1",
            "ACCENT_BORDER_CALLOUT_2",
            "ACCENT_BORDER_CALLOUT_3",
            "LEFT_ARROW_CALLOUT",
            "RIGHT_ARROW_CALLOUT",
            "UP_ARROW_CALLOUT",
            "DOWN_ARROW_CALLOUT",
            "LEFT_RIGHT_ARROW_CALLOUT",
            "UP_DOWN_ARROW_CALLOUT",
            "QUAD_ARROW_CALLOUT",
            "RECT",
            "ROUND_RECT",
            "ELLIPSE",
            "RIGHT_ARROW",
            "LEFT_ARROW",
            "UP_ARROW",
            "DOWN_ARROW",
            "LEFT_RIGHT_ARROW",
            "CURVED_LEFT_ARROW",
            "CURVED_RIGHT_ARROW",
            "CURVED_UP_ARROW",
            "CURVED_DOWN_ARROW",
            "BENT_ARROW",
            "BENT_UP_ARROW",
            "CIRCULAR_ARROW",
            "LEFT_CIRCULAR_ARROW",
            "LEFT_RIGHT_CIRCULAR_ARROW",
            "LEFT_RIGHT_UP_ARROW",
            "LEFT_UP_ARROW",
            "NOTCHED_RIGHT_ARROW",
            "QUAD_ARROW",
            "STRIPED_RIGHT_ARROW",
            "SWOOSH_ARROW",
            "UP_DOWN_ARROW",
            "UTURN_ARROW",
            "RIBBON",
            "RIBBON_2",
            "ELLIPSE_RIBBON",
            "ELLIPSE_RIBBON_2",
            "LEFT_RIGHT_RIBBON",
            "VERTICAL_SCROLL",
            "HORIZONTAL_SCROLL",
            "FLOW_CHART_PROCESS",
            "FLOW_CHART_PREDEFINED_PROCESS",
            "FLOW_CHART_ALTERNATE_PROCESS",
            "FLOW_CHART_CONNECTOR",
            "FLOW_CHART_DECISION",
            "FLOW_CHART_DOCUMENT",
            "FLOW_CHART_MULTIDOCUMENT",
            "FLOW_CHART_MAGNETIC_DISK",
            "FLOW_CHART_OFFPAGE_CONNECTOR",
            "FLOW_CHART_TERMINATOR",
            "FLOW_CHART_PREPARATION",
            "FLOW_CHART_MANUAL_INPUT",
            "FLOW_CHART_MANUAL_OPERATION",
            "FLOW_CHART_PUNCHED_CARD",
            "FLOW_CHART_PUNCHED_TAPE",
            "FLOW_CHART_SUMMING_JUNCTION",
            "FLOW_CHART_OR",
            "FLOW_CHART_COLLATE",
            "FLOW_CHART_SORT",
            "FLOW_CHART_EXTRACT",
            "FLOW_CHART_MERGE",
            "FLOW_CHART_OFFLINE_STORAGE",
            "FLOW_CHART_ONLINE_STORAGE",
            "FLOW_CHART_MAGNETIC_TAPE",
            "FLOW_CHART_MAGNETIC_DRUM",
            "FLOW_CHART_DISPLAY",
            "FLOW_CHART_DELAY",
            "FLOW_CHART_INPUT_OUTPUT",
            "FLOW_CHART_INTERNAL_STORAGE",
            "FOLDED_CORNER",
            "HOME_PLATE",
            "STAR_5",
            "IRREGULAR_SEAL_2",
            "BORDER_CALLOUT_1",
            "BORDER_CALLOUT_2",
            "BORDER_CALLOUT_3",
            "WEDGE_RECT_CALLOUT",
            "WEDGE_ROUND_RECT_CALLOUT",
            "WEDGE_ELLIPSE_CALLOUT",
            "CLOUD_CALLOUT");

    /** Open or line-like presets that look clearer without a fill. */
    private static final Set<String> NO_FILL = Set.of(
            "LINE",
            "LINE_INV",
            "ARC",
            "LEFT_BRACKET",
            "RIGHT_BRACKET",
            "LEFT_BRACE",
            "RIGHT_BRACE",
            "BRACKET_PAIR",
            "BRACE_PAIR",
            "FRAME",
            "HALF_FRAME",
            "CORNER");

    private static final Map<String, Category> CATEGORIES = categories();

    private UnsupportedShapesWorkbookGenerator() {
    }

    public record ShapeEntry(int type, String constant, String preset) {
    }

    private record Category(String sheetName, String label, String description,
                            int fill, int line) {
    }

    public static void main(String[] args) throws Exception {
        generate(DEFAULT_OUTPUT);
    }

    public static void generate(Path output) throws Exception {
        Path parent = output.toAbsolutePath().getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFCellStyle title = titleStyle(workbook);
            XSSFCellStyle subtitle = subtitleStyle(workbook);
            XSSFCellStyle header = headerStyle(workbook);
            XSSFCellStyle body = bodyStyle(workbook);
            XSSFCellStyle label = labelStyle(workbook);

            createIndexSheet(workbook, title, subtitle, header, body);
            int sheetIndex = 1;
            long nextId = 1;
            for (Map.Entry<String, Category> category : CATEGORIES.entrySet()) {
                List<ShapeEntry> entries = groupedShapes().get(category.getKey());
                nextId = createCategorySheet(workbook, sheetIndex++, category.getValue(),
                        entries, nextId, title, subtitle, label);
            }
            workbook.write(Files.newOutputStream(output));
        }
        System.out.println("generated " + output.toAbsolutePath()
                + " (" + unsupportedShapes().size() + " presets)");
    }

    /**
     * Returns every POI shape constant that is not supported by the renderer,
     * excluding connector presets. The list is sorted by constant name.
     */
    public static List<ShapeEntry> unsupportedShapes() {
        List<ShapeEntry> result = new ArrayList<>();
        for (Field field : ShapeTypes.class.getFields()) {
            if (field.getType() != int.class) {
                continue;
            }
            String constant = field.getName();
            if (SUPPORTED.contains(constant) || constant.contains("CONNECTOR")) {
                continue;
            }
            try {
                int type = field.getInt(null);
                String preset = STShapeType.Enum.forInt(type).toString();
                result.add(new ShapeEntry(type, constant, preset));
            } catch (ReflectiveOperationException | RuntimeException e) {
                throw new IllegalStateException("cannot resolve shape constant " + constant, e);
            }
        }
        result.sort(Comparator.comparing(ShapeEntry::constant));
        return result;
    }

    public static Map<String, List<ShapeEntry>> groupedShapes() {
        Map<String, List<ShapeEntry>> result = new LinkedHashMap<>();
        for (String key : CATEGORIES.keySet()) {
            result.put(key, new ArrayList<>());
        }
        for (ShapeEntry entry : unsupportedShapes()) {
            result.get(categoryKey(entry.constant())).add(entry);
        }
        return result;
    }

    private static long createCategorySheet(XSSFWorkbook workbook, int sheetIndex,
                                           Category category, List<ShapeEntry> entries,
                                           long nextId, XSSFCellStyle titleStyle,
                                           XSSFCellStyle subtitleStyle,
                                           XSSFCellStyle labelStyle) {
        Sheet sheet = workbook.createSheet(category.sheetName());
        sheet.setDisplayGridlines(false);
        for (int c = 0; c < TILES_PER_ROW * 2; c++) {
            sheet.setColumnWidth(c, 14 * 256);
        }
        title(sheet, 0, TILES_PER_ROW * 2,
                category.sheetName() + " — " + entries.size() + " unsupported presets",
                titleStyle);
        subtitle(sheet, 1, TILES_PER_ROW * 2,
                "Each preset is approximated by its bounding rectangle "
                        + "(SHAPE_APPROXIMATED). Labels show the POI constant and OOXML preset name.",
                subtitleStyle);
        Row spacer = sheet.createRow(2);
        spacer.setHeightInPoints(8);

        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
        int pageCount = pages(entries.size());
        for (int p = 0; p < pageCount; p++) {
            int base = FIRST_TILE_ROW + p * 21;
            for (int local = 0; local < 20; local++) {
                Row row = sheet.createRow(base + local);
                row.setHeightInPoints(local % TILE_ROWS == TILE_ROWS - 1 ? 30 : 20);
            }
            if (p < pageCount - 1) {
                Row pageSpacer = sheet.createRow(base + 20);
                pageSpacer.setHeightInPoints(8);
            }
        }

        for (int i = 0; i < entries.size(); i++) {
            ShapeEntry entry = entries.get(i);
            int col = (i % TILES_PER_ROW) * 2;
            int row = FIRST_TILE_ROW + (i / TILES_PER_PAGE) * 21
                    + ((i % TILES_PER_PAGE) / TILES_PER_ROW) * TILE_ROWS;
            XSSFSimpleShape shape = drawing.createSimpleShape(
                    new XSSFClientAnchor(4 * 9525, 3 * 9525, 4 * 9525, 3 * 9525,
                            col, row, col + 2, row + 3));
            shape.setShapeType(entry.type());
            if (NO_FILL.contains(entry.constant())) {
                shape.setNoFill(true);
            } else {
                shape.setFillColor((category.fill() >> 16) & 0xFF,
                        (category.fill() >> 8) & 0xFF, category.fill() & 0xFF);
            }
            shape.setLineStyleColor((category.line() >> 16) & 0xFF,
                    (category.line() >> 8) & 0xFF, category.line() & 0xFF);
            shape.setLineWidth(1.25);
            identify(shape, entry.constant(), nextId++);

            Row labelRow = sheet.getRow(row + 3);
            Cell label = labelRow.createCell(col);
            label.setCellValue(entry.constant() + "\n" + entry.preset());
            label.setCellStyle(labelStyle);
            sheet.addMergedRegion(new CellRangeAddress(row + 3, row + 3, col, col + 1));
        }

        sheet.setRepeatingRows(CellRangeAddress.valueOf("1:2"));
        for (int p = 1; p < pageCount; p++) {
            sheet.setRowBreak(FIRST_TILE_ROW + p * 21 - 1);
        }

        var setup = sheet.getPrintSetup();
        setup.setLandscape(true);
        setup.setPaperSize(org.apache.poi.ss.usermodel.PrintSetup.A4_PAPERSIZE);
        setup.setFitWidth((short) 1);
        setup.setFitHeight((short) 0);
        sheet.setFitToPage(true);
        sheet.setHorizontallyCenter(true);
        sheet.setMargin(org.apache.poi.ss.usermodel.PageMargin.LEFT, 0.4);
        sheet.setMargin(org.apache.poi.ss.usermodel.PageMargin.RIGHT, 0.4);
        sheet.setMargin(org.apache.poi.ss.usermodel.PageMargin.TOP, 0.4);
        sheet.setMargin(org.apache.poi.ss.usermodel.PageMargin.BOTTOM, 0.4);
        int lastRow = FIRST_TILE_ROW + pageCount * 20 + Math.max(0, pageCount - 1);
        workbook.setPrintArea(sheetIndex, "A1:J" + lastRow);
        return nextId;
    }

    private static void createIndexSheet(XSSFWorkbook workbook, XSSFCellStyle title,
                                         XSSFCellStyle subtitle, XSSFCellStyle header,
                                         XSSFCellStyle body) {
        Sheet sheet = workbook.createSheet("00 Index");
        sheet.setDisplayGridlines(false);
        sheet.setColumnWidth(0, 26 * 256);
        sheet.setColumnWidth(1, 10 * 256);
        sheet.setColumnWidth(2, 10 * 256);
        sheet.setColumnWidth(3, 62 * 256);

        title(sheet, 0, 4, "Unsupported shape presets", title);
        subtitle(sheet, 1, 4,
                "ShapeRenderer draws these presets as their bounding rectangle. "
                        + "Connectors and custom geometry are not included.",
                subtitle);
        Row headerRow = sheet.createRow(3);
        String[] headers = {"Sheet", "Presets", "Pages", "Contents"};
        for (int c = 0; c < headers.length; c++) {
            Cell cell = headerRow.createCell(c);
            cell.setCellValue(headers[c]);
            cell.setCellStyle(header);
        }
        int rowIndex = 4;
        int totalPages = 0;
        for (Map.Entry<String, Category> category : CATEGORIES.entrySet()) {
            List<ShapeEntry> entries = groupedShapes().get(category.getKey());
            int pages = pages(entries.size());
            totalPages += pages;
            Row row = sheet.createRow(rowIndex++);
            value(row, 0, category.getValue().sheetName(), body);
            value(row, 1, Integer.toString(entries.size()), body);
            value(row, 2, Integer.toString(pages), body);
            value(row, 3, category.getValue().description(), body);
        }
        Row totalRow = sheet.createRow(rowIndex);
        value(totalRow, 0, "Total", header);
        value(totalRow, 1, Integer.toString(unsupportedShapes().size()), header);
        value(totalRow, 2, Integer.toString(totalPages + 1), header);
        value(totalRow, 3, "Each shape = one SHAPE_APPROXIMATED warning (incl. index).", header);

        Row note = sheet.createRow(rowIndex + 2);
        Cell noteCell = note.createCell(0);
        noteCell.setCellValue("Excel で開くと実際の図形が表示されます。"
                + "xlsx2pdf4j の PDF ではすべて外接矩形で近似されます。");
        noteCell.setCellStyle(body);
        sheet.addMergedRegion(new CellRangeAddress(rowIndex + 2, rowIndex + 2, 0, 3));
        note.setHeightInPoints(28);

        sheet.createRow(rowIndex + 3).createCell(0).setCellValue(
                "Supported presets (rect, roundRect, ellipse, basic arrows, "
                        + "some flowchart shapes, etc.) and connectors are excluded.");
        Cell supportedNote = sheet.getRow(rowIndex + 3).getCell(0);
        supportedNote.setCellStyle(body);
        sheet.addMergedRegion(new CellRangeAddress(rowIndex + 3, rowIndex + 3, 0, 3));
        sheet.getRow(rowIndex + 3).setHeightInPoints(28);

        var setup = sheet.getPrintSetup();
        setup.setLandscape(true);
        setup.setFitWidth((short) 1);
        setup.setFitHeight((short) 1);
        sheet.setFitToPage(true);
        sheet.setMargin(org.apache.poi.ss.usermodel.PageMargin.LEFT, 0.5);
        sheet.setMargin(org.apache.poi.ss.usermodel.PageMargin.RIGHT, 0.5);
        sheet.setMargin(org.apache.poi.ss.usermodel.PageMargin.TOP, 0.5);
        sheet.setMargin(org.apache.poi.ss.usermodel.PageMargin.BOTTOM, 0.5);
        workbook.setPrintArea(0, "A1:D" + (rowIndex + 4));
    }

    private static String categoryKey(String constant) {
        if (constant.startsWith("ACTION_BUTTON_")) {
            return "action";
        }
        if (constant.contains("ARROW")) {
            return "arrows";
        }
        return "basic";
    }

    private static Map<String, Category> categories() {
        Map<String, Category> result = new LinkedHashMap<>();
        result.put("action", new Category("01 Action Buttons", "Action buttons",
                "Action button presets", 0xF2F2F2, 0x404040));
        return result;
    }

    private static int pages(int shapes) {
        return (shapes + TILES_PER_PAGE - 1) / TILES_PER_PAGE;
    }

    private static void identify(XSSFSimpleShape shape, String name, long id) {
        var nonVisual = shape.getCTShape().getNvSpPr().getCNvPr();
        nonVisual.setId(id);
        nonVisual.setName(name);
    }

    private static void value(Row row, int column, String text, CellStyle style) {
        Cell cell = row.createCell(column);
        cell.setCellValue(text);
        cell.setCellStyle(style);
    }

    private static void title(Sheet sheet, int row, int columns, String text,
                              XSSFCellStyle style) {
        Row r = sheet.getRow(row);
        if (r == null) {
            r = sheet.createRow(row);
        }
        Cell cell = r.createCell(0);
        cell.setCellValue(text);
        cell.setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(row, row, 0, columns - 1));
        r.setHeightInPoints(28);
    }

    private static void subtitle(Sheet sheet, int row, int columns, String text,
                                 XSSFCellStyle style) {
        Row r = sheet.getRow(row);
        if (r == null) {
            r = sheet.createRow(row);
        }
        Cell cell = r.createCell(0);
        cell.setCellValue(text);
        cell.setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(row, row, 0, columns - 1));
        r.setHeightInPoints(30);
    }

    private static XSSFFont font(XSSFWorkbook workbook, double size, boolean bold, int rgb) {
        XSSFFont font = workbook.createFont();
        font.setFontName("Meiryo");
        font.setFontHeightInPoints((short) size);
        font.setBold(bold);
        font.setColor(new XSSFColor(new byte[]{
                (byte) ((rgb >> 16) & 0xFF),
                (byte) ((rgb >> 8) & 0xFF),
                (byte) (rgb & 0xFF)}, null));
        return font;
    }

    private static XSSFCellStyle baseStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font(workbook, 10, false, 0x000000));
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private static XSSFCellStyle titleStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setFont(font(workbook, 16, true, 0xFFFFFF));
        style.setFillForegroundColor(new XSSFColor(new byte[]{0x1F, 0x49, 0x7D}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static XSSFCellStyle subtitleStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setFont(font(workbook, 9, false, 0x404040));
        style.setWrapText(true);
        style.setAlignment(HorizontalAlignment.LEFT);
        return style;
    }

    private static XSSFCellStyle headerStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setFont(font(workbook, 10, true, 0x1F497D));
        style.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0xDD, (byte) 0xEB, (byte) 0xF7}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        return style;
    }

    private static XSSFCellStyle bodyStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setWrapText(true);
        style.setBorderBottom(BorderStyle.HAIR);
        return style;
    }

    private static XSSFCellStyle labelStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setFont(font(workbook, 8, false, 0x404040));
        style.setWrapText(true);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.HAIR);
        return style;
    }
}
