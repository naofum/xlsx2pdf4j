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
import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFRadarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.TextAlign;
import org.apache.poi.xssf.usermodel.TextAutofit;
import org.apache.poi.xssf.usermodel.TextDirection;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFChildAnchor;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFConnector;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFShapeGroup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextBox;
import org.apache.poi.xssf.usermodel.XSSFTextParagraph;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.main.CTCustomGeometry2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGradientFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGradientStop;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.STCompoundLine;
import org.openxmlformats.schemas.drawingml.x2006.main.STLineEndType;
import org.openxmlformats.schemas.drawingml.x2006.main.STPresetLineDashVal;
import org.openxmlformats.schemas.drawingml.x2006.main.STTextAnchoringType;

import javax.imageio.ImageIO;
import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.GradientPaint;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/**
 * Builds a synthetic workbook that exercises the supported cell, shape and
 * print features. Run this class from the test classpath after test-compile.
 */
public final class ValidationWorkbookGenerator {

    private static final Path DEFAULT_OUTPUT = Path.of("testdata", "validation", "attributes.xlsx");
    private static long nextId = 1;

    private ValidationWorkbookGenerator() {
    }

    public static void main(String[] args) throws Exception {
        generate(DEFAULT_OUTPUT);
    }

    /**
     * Writes the validation workbook to {@code output}. The temporary
     * pre-gradient workbook is removed afterwards.
     */
    public static void generate(Path output) throws Exception {
        nextId = 1;
        Path parent = output.toAbsolutePath().getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        Path base = output.resolveSibling(output.getFileName().toString()
                .replaceFirst("\\.xlsx$", "") + "-base.xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            createCellsSheet(workbook);
            createTextSheet(workbook);
            createShapesSheet(workbook);
            createConnectorsSheet(workbook);
            createImagesSheet(workbook);
            createPrintSheet(workbook);
            createWarningsSheet(workbook);
            workbook.createSheet("99 Empty");
            try (OutputStream out = Files.newOutputStream(base)) {
                workbook.write(out);
            }
        }
        replaceCellGradient(base, output);
        Files.deleteIfExists(base);
        System.out.println("generated " + output.toAbsolutePath());
    }

    private static void createCellsSheet(XSSFWorkbook workbook) {
        Sheet sheet = workbook.createSheet("01 Cells");
        widths(sheet, 30, 40, 22, 22, 22, 22);
        XSSFCellStyle title = titleStyle(workbook);
        XSSFCellStyle section = sectionStyle(workbook);
        XSSFCellStyle label = labelStyle(workbook);
        XSSFCellStyle plain = baseStyle(workbook);
        int row = 0;
        title(sheet, row++, 6, "Cell attributes: fonts, fills, borders, formats, alignment, merge, rich text", title);

        row = section(sheet, row, "Fonts and colors", section);
        row = value(sheet, row, "Bold / italic / underline / strike", "Bold Italic Underline Strike",
                fontStyle(workbook, 12, true, true, true, true, 0x1F497D), label);
        row = value(sheet, row, "Font size 18 / color", "Large colored text",
                fontStyle(workbook, 18, false, false, false, false, 0xC00000), label);
        row = value(sheet, row, "Theme-like colors", "Blue / Red / Green",
                fontStyle(workbook, 12, false, false, false, false, 0x0070C0), label);

        row = section(sheet, row, "Fills", section);
        row = value(sheet, row, "Solid fill", "Solid fill", fillStyle(workbook, 0xDDEBF7), label);
        row = value(sheet, row, "Cell gradient fill", "Gradient fill", gradientCellStyle(workbook), label);
        row = value(sheet, row, "Pattern fill", "Pattern fill", patternStyle(workbook), label);

        row = section(sheet, row, "Borders", section);
        row = value(sheet, row, "Thin / medium / thick", "Borders", borderStyle(workbook, BorderStyle.THIN), label);
        row = value(sheet, row, "Dashed / dotted / double", "Borders", borderStyle(workbook, BorderStyle.DOUBLE), label);
        row = value(sheet, row, "Hair border", "Hair", borderStyle(workbook, BorderStyle.HAIR), label);

        row = section(sheet, row, "Number formats", section);
        row = number(sheet, row, "0.00 / #,##0.000", 1234.5678, numberStyle(workbook, "0.00"), label, plain);
        row = number(sheet, row, "Currency", 12345.67, numberStyle(workbook, "#,##0.00 [$¥-411]"), label, plain);
        row = number(sheet, row, "Percent", 0.1234, numberStyle(workbook, "0.00%"), label, plain);
        row = number(sheet, row, "Scientific", 12345.678, numberStyle(workbook, "0.00E+00"), label, plain);
        row = number(sheet, row, "Date", 45200, numberStyle(workbook, "yyyy/m/d"), label, plain);
        row = number(sheet, row, "Time", 0.625, numberStyle(workbook, "h:mm:ss"), label, plain);

        row = section(sheet, row, "Cell values", section);
        row = value(sheet, row, "Boolean", "TRUE", plain, label);
        sheet.getRow(row - 1).getCell(1).setCellValue(true);
        row = value(sheet, row, "Error", "#DIV/0!", plain, label);
        sheet.getRow(row - 1).getCell(1).setCellErrorValue(org.apache.poi.ss.usermodel.FormulaError.DIV0.getCode());
        row = value(sheet, row, "Formula with cache", "3", plain, label);
        XSSFCell formula = (XSSFCell) sheet.getRow(row - 1).getCell(1);
        formula.setCellFormula("1+2");
        row = value(sheet, row, "Hyperlink", "OpenAI", hyperlinkStyle(workbook), label);
        XSSFHyperlink hyperlink = (XSSFHyperlink) workbook.getCreationHelper().createHyperlink(
                org.apache.poi.common.usermodel.HyperlinkType.URL);
        hyperlink.setAddress("https://openai.com/");
        sheet.getRow(row - 1).getCell(1).setHyperlink(hyperlink);

        row = section(sheet, row, "Alignment and layout", section);
        row = value(sheet, row, "Center / middle", "Centered", alignmentStyle(workbook,
                HorizontalAlignment.CENTER, VerticalAlignment.CENTER), label);
        row = value(sheet, row, "Right / bottom", "Right", alignmentStyle(workbook,
                HorizontalAlignment.RIGHT, VerticalAlignment.BOTTOM), label);
        row = value(sheet, row, "Wrap text", "This text is wrapped inside a narrow cell and continues on another line",
                wrapStyle(workbook), label);
        sheet.getRow(row - 1).setHeightInPoints(60);
        row = value(sheet, row, "Shrink to fit", "This text shrinks to fit the cell width",
                shrinkStyle(workbook), label);
        row = value(sheet, row, "Indent", "Indented", indentStyle(workbook), label);

        row = section(sheet, row, "Merged cells", section);
        int merged = row;
        sheet.createRow(merged).createCell(0).setCellValue("Merged A:D");
        sheet.addMergedRegion(new CellRangeAddress(merged, merged + 1, 0, 3));
        sheet.getRow(merged).getCell(0).setCellStyle(fillStyle(workbook, 0xE2F0D9));
        row += 2;
        row = value(sheet, row, "Rich text", "Rich text", plain, label);
        XSSFRichTextString rich = new XSSFRichTextString();
        rich.append("Red ", font(workbook, 12, true, false, false, false, 0xC00000));
        rich.append("blue ", font(workbook, 12, false, true, false, false, 0x0070C0));
        rich.append("green", font(workbook, 12, false, false, true, false, 0x00B050));
        sheet.getRow(row - 1).getCell(1).setCellValue(rich);

        row = section(sheet, row, "Hidden and empty cells", section);
        row = value(sheet, row, "Hidden row below", "This row is hidden", plain, label);
        sheet.getRow(row - 1).setZeroHeight(true);
        row = value(sheet, row, "Empty styled cell", "", fillStyle(workbook, 0xFFF2CC), label);
        row = value(sheet, row, "Very long text", "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", plain, label);

        sheet.createFreezePane(1, 1);
        sheet.getPrintSetup().setLandscape(true);
        sheet.setFitToPage(true);
        sheet.getPrintSetup().setFitWidth((short) 1);
        sheet.getPrintSetup().setFitHeight((short) 1);
        workbook.setPrintArea(0, "A1:F" + row);
    }

    private static void createTextSheet(XSSFWorkbook workbook) {
        Sheet sheet = workbook.createSheet("02 Text");
        widths(sheet, 30, 55, 15, 15);
        XSSFCellStyle title = titleStyle(workbook);
        XSSFCellStyle label = labelStyle(workbook);
        XSSFCellStyle plain = baseStyle(workbook);
        int row = 0;
        title(sheet, row++, 4, "Text: rich text, wrapping, rotation, shrink, vertical alignment", title);

        row = value(sheet, row, "Multi-font rich text", "", plain, label);
        XSSFRichTextString rich = new XSSFRichTextString();
        rich.append("Bold ", font(workbook, 13, true, false, false, false, 0x000000));
        rich.append("Italic ", font(workbook, 13, false, true, false, false, 0x7030A0));
        rich.append("Underline ", font(workbook, 13, false, false, true, false, 0x0070C0));
        rich.append("Strike", font(workbook, 13, false, false, false, true, 0xC00000));
        sheet.getRow(row - 1).getCell(1).setCellValue(rich);
        sheet.getRow(row - 1).setHeightInPoints(30);

        row = value(sheet, row, "Multiline text", "Line 1\nLine 2\nLine 3\nLine 4",
                wrapStyle(workbook), label);
        sheet.getRow(row - 1).setHeightInPoints(80);
        row = value(sheet, row, "Japanese / symbols", "日本語の表現、①②③、→←↑↓、℃、π、①㈱",
                fontStyle(workbook, 14, false, false, false, false, 0x000000), label);
        sheet.getRow(row - 1).setHeightInPoints(28);
        row = value(sheet, row, "Wide characters", "ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰ",
                fontStyle(workbook, 14, false, false, false, false, 0x000000), label);
        row = value(sheet, row, "Wrap + indent", "Wrapped and indented text that spans multiple lines in this cell",
                wrapIndentStyle(workbook), label);
        sheet.getRow(row - 1).setHeightInPoints(60);
        row = value(sheet, row, "Shrink to fit", "This text should shrink to fit in the available cell width",
                shrinkStyle(workbook), label);
        row = rotationExample(sheet, row, "45°", "Rotation 45°", 45, 60, label, workbook);
        row = rotationExample(sheet, row, "90°", "Rotation 90°", 90, 100, label, workbook);
        row = value(sheet, row, "Vertical bottom", "Bottom aligned", alignmentStyle(workbook,
                HorizontalAlignment.LEFT, VerticalAlignment.BOTTOM), label);
        sheet.getRow(row - 1).setHeightInPoints(45);

        int merged = row;
        sheet.createRow(merged).createCell(0).setCellValue("Merged multiline");
        sheet.addMergedRegion(new CellRangeAddress(merged, merged + 3, 0, 2));
        Cell mergedCell = sheet.getRow(merged).getCell(0);
        mergedCell.setCellValue("Merged region with\nwrapped text and\ncenter alignment");
        mergedCell.setCellStyle(wrapCenterStyle(workbook));

        sheet.getPrintSetup().setLandscape(true);
        sheet.setFitToPage(true);
        sheet.getPrintSetup().setFitWidth((short) 1);
        sheet.getPrintSetup().setFitHeight((short) 1);
        workbook.setPrintArea(1, "A1:D" + (merged + 5));
    }

    private static void createShapesSheet(XSSFWorkbook workbook) {
        Sheet sheet = workbook.createSheet("03 Shapes");
        widths(sheet, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12);
        for (int r = 0; r < 34; r++) {
            sheet.createRow(r).setHeightInPoints(20);
        }
        XSSFCellStyle title = titleStyle(workbook);
        title(sheet, 0, 12, "Shapes: geometry, fill, line, arrows, text, autofit, group, custom geometry", title);
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();

        int[][] boxes = {
                {1, 2, 4, 5}, {5, 2, 8, 5}, {9, 2, 12, 5},
                {1, 7, 4, 10}, {5, 7, 8, 10}, {9, 7, 12, 10},
                {1, 12, 4, 15}, {5, 12, 8, 15}, {9, 12, 12, 15},
                {1, 17, 4, 20}, {5, 17, 8, 20}, {9, 17, 12, 20},
                {1, 22, 4, 25}, {5, 22, 8, 25}, {9, 22, 12, 25},
                {1, 27, 4, 30}, {5, 27, 8, 30}, {9, 27, 12, 30}
        };
        int[] types = {
                ShapeTypes.RECT, ShapeTypes.ROUND_RECT, ShapeTypes.ELLIPSE,
                ShapeTypes.RIGHT_ARROW, ShapeTypes.LEFT_RIGHT_ARROW, ShapeTypes.DOWN_ARROW,
                ShapeTypes.FLOW_CHART_PROCESS, ShapeTypes.FLOW_CHART_DECISION, ShapeTypes.FLOW_CHART_DOCUMENT,
                ShapeTypes.FLOW_CHART_MULTIDOCUMENT, ShapeTypes.FLOW_CHART_MAGNETIC_DISK,
                ShapeTypes.FLOW_CHART_OFFPAGE_CONNECTOR,
                ShapeTypes.FOLDED_CORNER, ShapeTypes.HOME_PLATE, ShapeTypes.STAR_5,
                ShapeTypes.IRREGULAR_SEAL_2, ShapeTypes.WEDGE_RECT_CALLOUT, ShapeTypes.BORDER_CALLOUT_2
        };
        String[] labels = {
                "Rect", "Round rect", "Ellipse",
                "Right arrow", "Both arrows", "Down arrow",
                "Process", "Decision", "Document",
                "Multidocument", "Magnetic disk", "Off-page",
                "Folded corner", "Home plate", "Star",
                "Seal", "Callout", "Border callout"
        };
        int[] colors = {
                0xDDEBF7, 0xE2F0D9, 0xFFF2CC,
                0xFCE4D6, 0xFF99CC, 0xD9E1F2,
                0xE2F0D9, 0xFFF2CC, 0xDDEBF7,
                0xE4DFEC, 0xE2F0D9, 0xFCE4D6,
                0xFFF2CC, 0xDDEBF7, 0xFFD966,
                0xE4DFEC, 0xFF99CC, 0xFCE4D6
        };
        for (int i = 0; i < boxes.length; i++) {
            XSSFSimpleShape shape = addShape(drawing, boxes[i], types[i], colors[i], labels[i]);
            if (types[i] == ShapeTypes.RIGHT_ARROW) {
                adjustment(shape, "adj1", "65000");
                adjustment(shape, "adj2", "75000");
            } else if (types[i] == ShapeTypes.LEFT_RIGHT_ARROW) {
                adjustment(shape, "adj1", "54870");
                adjustment(shape, "adj2", "67211");
            } else if (types[i] == ShapeTypes.DOWN_ARROW) {
                adjustment(shape, "adj1", "60000");
                adjustment(shape, "adj2", "70000");
            } else if (types[i] == ShapeTypes.WEDGE_RECT_CALLOUT) {
                adjustment(shape, "adj1", "-20000");
                adjustment(shape, "adj2", "60000");
            }
        }

        XSSFSimpleShape gradient = drawing.createSimpleShape(anchor(1, 31, 4, 34));
        identify(gradient, "Gradient shape");
        gradient.setShapeType(ShapeTypes.ROUND_RECT);
        gradient.setLineStyleColor(0x1F, 0x49, 0x7D);
        gradient.setLineWidth(1.5);
        gradientFill(gradient, 0x4F81BD, 0xC6E0B4);
        gradient.setText("Gradient");
        centerAll(gradient, 0x000000, 11);
        adjustment(gradient, "adj", "30000");

        XSSFSimpleShape shadow = drawing.createSimpleShape(anchor(5, 31, 8, 34));
        identify(shadow, "Shadow shape");
        shadow.setShapeType(ShapeTypes.RECT);
        shadow.setFillColor(0xFF, 0xC0, 0x00);
        shadow.setLineStyleColor(0x00, 0x00, 0x00);
        shadow.setText("Shadow");
        centerAll(shadow, 0x000000, 11);
        shadow(shadow);

        XSSFSimpleShape doubleLine = drawing.createSimpleShape(anchor(9, 31, 12, 34));
        identify(doubleLine, "Double line shape");
        doubleLine.setShapeType(ShapeTypes.RECT);
        doubleLine.setFillColor(0xF2, 0xF2, 0xF2);
        doubleLine.setLineStyleColor(0xC0, 0x00, 0x00);
        doubleLine.setLineWidth(6);
        doubleLine.getCTShape().getSpPr().getLn().setCmpd(STCompoundLine.DBL);
        doubleLine.setText("Double line");
        centerAll(doubleLine, 0xC00000, 12);

        XSSFSimpleShape multi = drawing.createSimpleShape(anchor(1, 36, 4, 40));
        identify(multi, "Paragraph shape");
        multi.setShapeType(ShapeTypes.RECT);
        multi.setFillColor(0xEE, 0xEE, 0xEE);
        multi.setLineStyleColor(0x00, 0x00, 0x00);
        multi.setWordWrap(true);
        multi.clearText();
        XSSFTextParagraph p1 = multi.addNewTextParagraph();
        addRun(p1, "Left", 10, 0x000000);
        p1.setTextAlign(TextAlign.LEFT);
        XSSFTextParagraph p2 = multi.addNewTextParagraph();
        addRun(p2, "Center", 10, 0xC00000);
        p2.setTextAlign(TextAlign.CENTER);
        XSSFTextParagraph p3 = multi.addNewTextParagraph();
        addRun(p3, "Right", 10, 0x0070C0);
        p3.setTextAlign(TextAlign.RIGHT);

        XSSFSimpleShape norm = drawing.createSimpleShape(anchor(5, 36, 8, 40));
        identify(norm, "Norm autofit shape");
        norm.setShapeType(ShapeTypes.RECT);
        norm.setFillColor(0xE2, 0xF0, 0xD9);
        norm.setLineStyleColor(0x54, 0x82, 0x35);
        norm.setWordWrap(true);
        norm.setTextAutofit(TextAutofit.NORMAL);
        norm.getCTShape().getTxBody().getBodyPr().getNormAutofit().setFontScale(65000);
        norm.setText("Norm autofit: this text is deliberately long enough to need shrinking");
        centerAll(norm, 0x000000, 15);

        XSSFSimpleShape anchorCtr = drawing.createSimpleShape(anchor(9, 36, 12, 40));
        identify(anchorCtr, "Anchor center shape");
        anchorCtr.setShapeType(ShapeTypes.RECT);
        anchorCtr.setFillColor(0xFC, 0xE4, 0xD6);
        anchorCtr.setLineStyleColor(0xC6, 0x59, 0x11);
        anchorCtr.getCTShape().getTxBody().getBodyPr().setAnchorCtr(true);
        anchorCtr.setText("Anchor center");
        centerAll(anchorCtr, 0x000000, 13);
        anchorCtr.getTextParagraphs().get(0).setTextAlign(TextAlign.LEFT);

        XSSFSimpleShape vertical = drawing.createSimpleShape(anchor(13, 1, 15, 8));
        identify(vertical, "Vertical text shape");
        vertical.setShapeType(ShapeTypes.RECT);
        vertical.setFillColor(0xD9, 0xE1, 0xF2);
        vertical.setLineStyleColor(0x1F, 0x49, 0x7D);
        vertical.setTextDirection(TextDirection.VERTICAL);
        vertical.setText("縦書きテキスト");
        centerAll(vertical, 0x000000, 12);

        XSSFSimpleShape custom = drawing.createSimpleShape(anchor(13, 10, 15, 14));
        identify(custom, "Custom geometry shape");
        custom.setShapeType(ShapeTypes.RECT);
        custom.setFillColor(0xA9, 0xD0, 0x8E);
        custom.setLineStyleColor(0x00, 0x00, 0x00);
        custom.setLineWidth(1.5);
        customGeometry(custom);

        XSSFSimpleShape rotated = drawing.createSimpleShape(anchor(13, 16, 15, 20));
        identify(rotated, "Rotated shape");
        rotated.setShapeType(ShapeTypes.RECT);
        rotated.setFillColor(0xFF, 0xD9, 0x66);
        rotated.setLineStyleColor(0x00, 0x00, 0x00);
        rotated.setText("Rotated / flipped");
        centerAll(rotated, 0x000000, 10);
        rotated.getCTShape().getSpPr().getXfrm().setRot(2700000);
        rotated.getCTShape().getSpPr().getXfrm().setFlipH(true);

        XSSFTextBox textBox = drawing.createTextbox(anchor(13, 22, 15, 26));
        identify(textBox, "Text box");
        textBox.setText("Text box\nno fill / no line");
        centerAll(textBox, 0x000000, 11);
        noFill(textBox);
        noLine(textBox);

        XSSFShapeGroup group = drawing.createGroup(anchor(13, 28, 15, 33));
        identify(group, "Group");
        XSSFSimpleShape child1 = group.createSimpleShape(new XSSFChildAnchor(0, 0, 450000, 450000));
        identify(child1, "Group rectangle");
        child1.setShapeType(ShapeTypes.RECT);
        child1.setFillColor(0xFF, 0x99, 0xCC);
        child1.setLineStyleColor(0x00, 0x00, 0x00);
        XSSFSimpleShape child2 = group.createSimpleShape(new XSSFChildAnchor(550000, 550000, 1000000, 1000000));
        identify(child2, "Group ellipse");
        child2.setShapeType(ShapeTypes.ELLIPSE);
        child2.setFillColor(0x99, 0xCC, 0xFF);
        child2.setLineStyleColor(0x00, 0x00, 0x00);

        sheet.getPrintSetup().setLandscape(true);
        sheet.setFitToPage(true);
        sheet.getPrintSetup().setFitWidth((short) 1);
        sheet.getPrintSetup().setFitHeight((short) 1);
        workbook.setPrintArea(2, "A1:P40");
    }

    private static void createConnectorsSheet(XSSFWorkbook workbook) {
        Sheet sheet = workbook.createSheet("04 Connectors");
        widths(sheet, 14, 14, 14, 14, 14, 14, 14, 14);
        for (int r = 0; r < 30; r++) {
            sheet.createRow(r).setHeightInPoints(20);
        }
        XSSFCellStyle title = titleStyle(workbook);
        title(sheet, 0, 8, "Connectors and line shapes", title);
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();

        XSSFSimpleShape a = addShape(drawing, new int[]{1, 2, 3, 4}, ShapeTypes.RECT, 0xDDEBF7, "A");
        XSSFSimpleShape b = addShape(drawing, new int[]{6, 2, 8, 4}, ShapeTypes.RECT, 0xE2F0D9, "B");
        XSSFSimpleShape c = addShape(drawing, new int[]{1, 7, 3, 9}, ShapeTypes.RECT, 0xFFF2CC, "C");
        XSSFSimpleShape d = addShape(drawing, new int[]{6, 8, 8, 10}, ShapeTypes.RECT, 0xFCE4D6, "D");
        XSSFSimpleShape e = addShape(drawing, new int[]{1, 12, 3, 14}, ShapeTypes.RECT, 0xE2F0D9, "E");
        XSSFSimpleShape f = addShape(drawing, new int[]{6, 13, 8, 15}, ShapeTypes.RECT, 0xE4DFEC, "F");
        XSSFSimpleShape g = addShape(drawing, new int[]{1, 17, 3, 19}, ShapeTypes.RECT, 0xFCE4D6, "G");
        XSSFSimpleShape h = addShape(drawing, new int[]{6, 18, 8, 20}, ShapeTypes.RECT, 0xDDEBF7, "H");

        connector(drawing, new XSSFClientAnchor(0, 0, 0, 0, 3, 3, 6, 3), a, b,
                ShapeTypes.STRAIGHT_CONNECTOR_1, 0x0070C0, 2.25, "triangle");
        connector(drawing, new XSSFClientAnchor(0, 0, 0, 0, 3, 9, 6, 9), c, d,
                ShapeTypes.BENT_CONNECTOR_2, 0xC00000, 1.5, "stealth");
        connector(drawing, new XSSFClientAnchor(0, 0, 0, 0, 3, 14, 6, 14), e, f,
                ShapeTypes.BENT_CONNECTOR_3, 0x548235, 1.5, "triangle");
        connector(drawing, new XSSFClientAnchor(0, 0, 0, 0, 3, 19, 6, 19), g, h,
                ShapeTypes.BENT_CONNECTOR_4, 0x7030A0, 1.5, "diamond");

        XSSFSimpleShape line = drawing.createSimpleShape(anchor(1, 23, 8, 23));
        identify(line, "Arrow line");
        line.setShapeType(ShapeTypes.LINE);
        line.setLineStyleColor(0xC0, 0x00, 0x00);
        line.setLineWidth(3);
        var lineLn = line.getCTShape().getSpPr().getLn();
        lineLn.addNewTailEnd().setType(STLineEndType.STEALTH);
        lineLn.addNewHeadEnd().setType(STLineEndType.OVAL);
        lineLn.addNewPrstDash().setVal(STPresetLineDashVal.LG_DASH);

        XSSFSimpleShape dashed = drawing.createSimpleShape(anchor(1, 25, 8, 25));
        identify(dashed, "Dashed line");
        dashed.setShapeType(ShapeTypes.LINE);
        dashed.setLineStyleColor(0x00, 0x70, 0xC0);
        dashed.setLineWidth(2);
        var dashedLn = dashed.getCTShape().getSpPr().getLn();
        dashedLn.addNewPrstDash().setVal(STPresetLineDashVal.DASH);
        dashedLn.addNewTailEnd().setType(STLineEndType.TRIANGLE);

        XSSFSimpleShape doubleLine = drawing.createSimpleShape(anchor(1, 27, 8, 27));
        identify(doubleLine, "Double line");
        doubleLine.setShapeType(ShapeTypes.LINE);
        doubleLine.setLineStyleColor(0x00, 0x00, 0x00);
        doubleLine.setLineWidth(6);
        doubleLine.getCTShape().getSpPr().getLn().setCmpd(STCompoundLine.DBL);
        doubleLine.getCTShape().getSpPr().getLn().addNewTailEnd().setType(STLineEndType.TRIANGLE);

        sheet.getPrintSetup().setLandscape(true);
        sheet.setFitToPage(true);
        sheet.getPrintSetup().setFitWidth((short) 1);
        sheet.getPrintSetup().setFitHeight((short) 1);
        workbook.setPrintArea(3, "A1:H29");
    }

    private static void createImagesSheet(XSSFWorkbook workbook) throws IOException {
        Sheet sheet = workbook.createSheet("05 Images");
        widths(sheet, 24, 24, 24, 24);
        for (int r = 0; r < 18; r++) {
            sheet.createRow(r).setHeightInPoints(20);
        }
        XSSFCellStyle title = titleStyle(workbook);
        title(sheet, 0, 4, "Images: PNG with alpha, JPEG, and positioning", title);
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();

        int png = workbook.addPicture(pngImage(), Workbook.PICTURE_TYPE_PNG);
        int jpeg = workbook.addPicture(jpegImage(), Workbook.PICTURE_TYPE_JPEG);
        XSSFPicture p1 = drawing.createPicture(new XSSFClientAnchor(0, 0, 0, 0, 1, 2, 4, 8), png);
        identify(p1, "transparent-png");
        p1.getCTPicture().getNvPicPr().getCNvPr().setName("transparent-png");
        XSSFPicture p2 = drawing.createPicture(new XSSFClientAnchor(0, 0, 0, 0, 6, 2, 9, 8), jpeg);
        identify(p2, "gradient-jpeg");
        p2.getCTPicture().getNvPicPr().getCNvPr().setName("gradient-jpeg");

        sheet.getPrintSetup().setLandscape(true);
        sheet.setFitToPage(true);
        sheet.getPrintSetup().setFitWidth((short) 1);
        sheet.getPrintSetup().setFitHeight((short) 1);
        workbook.setPrintArea(4, "A1:J16");
    }

    private static void createPrintSheet(XSSFWorkbook workbook) {
        Sheet sheet = workbook.createSheet("06 Print");
        widths(sheet, 10, 24, 12, 12, 14, 24);
        XSSFCellStyle title = titleStyle(workbook);
        XSSFCellStyle header = sectionStyle(workbook);
        XSSFCellStyle plain = baseStyle(workbook);
        title(sheet, 0, 6, "Print settings", title);
        Row headerRow = sheet.createRow(1);
        String[] headers = {"No.", "Item", "Qty", "Price", "Amount", "Note"};
        for (int c = 0; c < headers.length; c++) {
            Cell cell = headerRow.createCell(c);
            cell.setCellValue(headers[c]);
            cell.setCellStyle(header);
        }
        for (int r = 2; r < 42; r++) {
            Row row = sheet.createRow(r);
            for (int c = 0; c < 6; c++) {
                Cell cell = row.createCell(c);
                cell.setCellStyle(plain);
                if (c == 0) cell.setCellValue(r - 1);
                else if (c == 1) cell.setCellValue("Validation item " + (r - 1));
                else if (c == 2) cell.setCellValue((r % 9) + 1);
                else if (c == 3) cell.setCellValue(100 + r * 3.5);
                else if (c == 4) cell.setCellFormula("C" + (r + 1) + "*D" + (r + 1));
                else cell.setCellValue("Page break / title test");
            }
            row.setHeightInPoints(18);
        }
        sheet.setRepeatingRows(CellRangeAddress.valueOf("1:2"));
        sheet.setRowBreak(18);
        sheet.setRowBreak(34);
        sheet.setAutobreaks(false);
        var setup = sheet.getPrintSetup();
        setup.setLandscape(true);
        setup.setPaperSize(PrintSetup.A4_PAPERSIZE);
        setup.setFitWidth((short) 1);
        setup.setFitHeight((short) 0);
        sheet.setFitToPage(true);
        sheet.setHorizontallyCenter(true);
        sheet.setVerticallyCenter(false);
        sheet.setMargin(PageMargin.LEFT, 0.5);
        sheet.setMargin(PageMargin.RIGHT, 0.5);
        sheet.setMargin(PageMargin.TOP, 0.6);
        sheet.setMargin(PageMargin.BOTTOM, 0.6);
        sheet.getHeader().setLeft("&\"Arial,Bold\"Validation Workbook");
        sheet.getHeader().setCenter("&P / &N");
        sheet.getHeader().setRight("&D &T");
        sheet.getFooter().setLeft("&F");
        sheet.getFooter().setCenter("xlsx2pdf4j validation");
        sheet.getFooter().setRight("Page &P of &N");
        workbook.setPrintArea(5, "A1:F42");
    }

    private static void createWarningsSheet(XSSFWorkbook workbook) throws IOException {
        Sheet sheet = workbook.createSheet("07 Warnings");
        widths(sheet, 28, 30, 20, 20);
        for (int r = 0; r < 16; r++) {
            sheet.createRow(r).setHeightInPoints(20);
        }
        XSSFCellStyle title = titleStyle(workbook);
        XSSFCellStyle label = labelStyle(workbook);
        title(sheet, 0, 7, "Warning cases: chart, unsupported image, uncached formula, approximated shape", title);
        sheet.getRow(0).createCell(7).setCellValue("Category");
        sheet.getRow(0).createCell(8).setCellValue("Value");
        for (int r = 1; r <= 4; r++) {
            sheet.getRow(r).createCell(7).setCellValue("C" + r);
            sheet.getRow(r).createCell(8).setCellValue(r);
        }
        sheet.getRow(1).createCell(0).setCellValue("Chart (unsupported radar)");
        sheet.getRow(1).getCell(0).setCellStyle(label);
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
        XSSFSheet xssfSheet = (XSSFSheet) sheet;
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                xssfSheet, new CellRangeAddress(1, 4, 7, 7));
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                xssfSheet, new CellRangeAddress(1, 4, 8, 8));
        XSSFChart chart = drawing.createChart(new XSSFClientAnchor(0, 0, 0, 0, 1, 2, 4, 10));
        XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
        XDDFRadarChartData radar = (XDDFRadarChartData) chart.createData(
                ChartTypes.RADAR, categoryAxis, valueAxis);
        radar.addSeries(categories, values).setTitle("Values", null);
        chart.plot(radar);
        chart.setTitleText("Radar (unsupported)");

        sheet.createRow(3).createCell(0).setCellValue("Unsupported image");
        sheet.getRow(3).getCell(0).setCellStyle(label);
        int emf = workbook.addPicture(new byte[]{1, 2, 3, 4, 5}, Workbook.PICTURE_TYPE_EMF);
        XSSFPicture picture = drawing.createPicture(new XSSFClientAnchor(0, 0, 0, 0, 5, 2, 7, 6), emf);
        identify(picture, "fake-emf");

        Row uncached = sheet.getRow(5);
        if (uncached == null) {
            uncached = sheet.createRow(5);
        }
        Cell formulaLabel = uncached.createCell(0);
        formulaLabel.setCellValue("Uncached formula");
        formulaLabel.setCellStyle(label);
        uncached.createCell(1).setCellFormula("FOO(1)");

        sheet.createRow(7).createCell(0).setCellValue("Approximated shape");
        sheet.getRow(7).getCell(0).setCellStyle(label);
        XSSFSimpleShape approximated = drawing.createSimpleShape(
                new XSSFClientAnchor(0, 0, 0, 0, 1, 8, 3, 12));
        identify(approximated, "Approximated action button");
        approximated.setShapeType(ShapeTypes.ACTION_BUTTON_BLANK);
        approximated.setFillColor(0xC0, 0x00, 0x00);

        sheet.getPrintSetup().setLandscape(true);
        sheet.setFitToPage(true);
        sheet.getPrintSetup().setFitWidth((short) 1);
        sheet.getPrintSetup().setFitHeight((short) 1);
        workbook.setPrintArea(6, "A1:G14");
    }

    private static XSSFSimpleShape addShape(XSSFDrawing drawing, int[] box, int type, int rgb, String text) {
        XSSFSimpleShape shape = drawing.createSimpleShape(anchor(box[0], box[1], box[2], box[3]));
        identify(shape, "Shape");
        shape.setShapeType(type);
        shape.setFillColor((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
        shape.setLineStyleColor(0x00, 0x00, 0x00);
        shape.setLineWidth(1.0);
        if (text != null && !text.isEmpty()) {
            shape.setText(text);
            centerAll(shape, 0x000000, 11);
        }
        return shape;
    }

    private static void identify(XSSFSimpleShape shape, String name) {
        long id = nextId++;
        var nonVisual = shape.getCTShape().getNvSpPr().getCNvPr();
        nonVisual.setId(id);
        nonVisual.setName(name + " " + id);
    }

    private static void identify(XSSFConnector connector, String name) {
        long id = nextId++;
        var nonVisual = connector.getCTConnector().getNvCxnSpPr().getCNvPr();
        nonVisual.setId(id);
        nonVisual.setName(name + " " + id);
    }

    private static void identify(XSSFPicture picture, String name) {
        long id = nextId++;
        var nonVisual = picture.getCTPicture().getNvPicPr().getCNvPr();
        nonVisual.setId(id);
        nonVisual.setName(name + " " + id);
    }

    private static void identify(XSSFShapeGroup group, String name) {
        long id = nextId++;
        var nonVisual = group.getCTGroupShape().getNvGrpSpPr().getCNvPr();
        nonVisual.setId(id);
        nonVisual.setName(name + " " + id);
    }

    private static XSSFClientAnchor anchor(int c1, int r1, int c2, int r2) {
        return new XSSFClientAnchor(0, 0, 0, 0, c1, r1, c2, r2);
    }

    private static void adjustment(XSSFSimpleShape shape, String name, String value) {
        var list = shape.getCTShape().getSpPr().getPrstGeom().getAvLst();
        var guide = list.addNewGd();
        guide.setName(name);
        guide.setFmla("val " + value);
    }

    private static void centerAll(XSSFSimpleShape shape, int color, double size) {
        var body = shape.getCTShape().getTxBody().getBodyPr();
        body.setAnchor(STTextAnchoringType.CTR);
        for (XSSFTextParagraph paragraph : shape.getTextParagraphs()) {
            paragraph.setTextAlign(TextAlign.CENTER);
            for (var run : paragraph.getTextRuns()) {
                run.setFontSize(size);
                run.setFontColor(new Color((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF));
            }
        }
    }

    private static void addRun(XSSFTextParagraph paragraph, String text, double size, int color) {
        var run = paragraph.addNewTextRun();
        run.setText(text);
        run.setFontSize(size);
        run.setFontColor(new Color((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF));
        run.setFontFamily("Meiryo", (byte) 0, (byte) 0, false);
    }

    private static void gradientFill(XSSFSimpleShape shape, int from, int to) {
        var spPr = shape.getCTShape().getSpPr();
        if (spPr.isSetSolidFill()) {
            spPr.unsetSolidFill();
        }
        if (spPr.isSetNoFill()) {
            spPr.unsetNoFill();
        }
        CTGradientFillProperties gradient = spPr.addNewGradFill();
        var stops = gradient.addNewGsLst();
        gradientStop(stops.addNewGs(), 0, from);
        gradientStop(stops.addNewGs(), 100000, to);
        gradient.addNewLin().setAng(5400000);
    }

    private static void gradientStop(CTGradientStop stop, int position, int rgb) {
        stop.setPos(position);
        stop.addNewSrgbClr().setVal(new byte[]{
                (byte) ((rgb >> 16) & 0xFF),
                (byte) ((rgb >> 8) & 0xFF),
                (byte) (rgb & 0xFF)});
    }

    private static void shadow(XSSFSimpleShape shape) {
        var spPr = shape.getCTShape().getSpPr();
        var effect = spPr.isSetEffectLst() ? spPr.getEffectLst() : spPr.addNewEffectLst();
        var shadow = effect.addNewOuterShdw();
        shadow.setBlurRad(40000);
        shadow.setDist(30000);
        shadow.setDir(2700000);
        shadow.addNewSrgbClr().setVal(new byte[]{(byte) 0x80, (byte) 0x80, (byte) 0x80});
    }

    private static void noFill(XSSFSimpleShape shape) {
        var spPr = shape.getCTShape().getSpPr();
        if (spPr.isSetSolidFill()) {
            spPr.unsetSolidFill();
        }
        if (spPr.isSetGradFill()) {
            spPr.unsetGradFill();
        }
        spPr.addNewNoFill();
    }

    private static void noLine(XSSFSimpleShape shape) {
        var spPr = shape.getCTShape().getSpPr();
        CTLineProperties line = spPr.isSetLn() ? spPr.getLn() : spPr.addNewLn();
        if (line.isSetSolidFill()) {
            line.unsetSolidFill();
        }
        line.addNewNoFill();
    }

    private static void customGeometry(XSSFSimpleShape shape) {
        var spPr = shape.getCTShape().getSpPr();
        spPr.unsetPrstGeom();
        CTCustomGeometry2D geom = spPr.addNewCustGeom();
        geom.addNewAvLst();
        geom.addNewGdLst();
        geom.addNewAhLst();
        geom.addNewCxnLst();
        var rect = geom.addNewRect();
        rect.setL(0L);
        rect.setT(0L);
        rect.setR(0L);
        rect.setB(0L);
        var pathList = geom.addNewPathLst();
        var path = pathList.addNewPath();
        path.setW(100);
        path.setH(100);
        moveTo(path, 0, 0);
        lineTo(path, 40, 0);
        lineTo(path, 40, 60);
        lineTo(path, 100, 60);
        lineTo(path, 100, 100);
        lineTo(path, 0, 100);
        path.addNewClose();
    }

    private static void moveTo(org.openxmlformats.schemas.drawingml.x2006.main.CTPath2D path,
                               int x, int y) {
        var point = path.addNewMoveTo().addNewPt();
        point.setX(x);
        point.setY(y);
    }

    private static void lineTo(org.openxmlformats.schemas.drawingml.x2006.main.CTPath2D path,
                               int x, int y) {
        var point = path.addNewLnTo().addNewPt();
        point.setX(x);
        point.setY(y);
    }

    private static void connector(XSSFDrawing drawing, XSSFClientAnchor anchor,
                                  XSSFSimpleShape start, XSSFSimpleShape end,
                                  int type, int rgb, double width, String endArrow) {
        XSSFConnector connector = drawing.createConnector(anchor);
        identify(connector, "Connector");
        connector.setShapeType(type);
        connector.setLineStyleColor((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
        connector.setLineWidth(width);
        var connectorNonVisual = connector.getCTConnector().getNvCxnSpPr();
        var properties = connectorNonVisual.getCNvCxnSpPr() != null
                ? connectorNonVisual.getCNvCxnSpPr()
                : connectorNonVisual.addNewCNvCxnSpPr();
        properties.addNewStCxn().setId(start.getCTShape().getNvSpPr().getCNvPr().getId());
        properties.getStCxn().setIdx(3);
        properties.addNewEndCxn().setId(end.getCTShape().getNvSpPr().getCNvPr().getId());
        properties.getEndCxn().setIdx(1);
        var line = connector.getCTConnector().getSpPr().getLn();
        var tail = line.addNewTailEnd();
        tail.setType(STLineEndType.Enum.forString(endArrow));
        tail.setW(org.openxmlformats.schemas.drawingml.x2006.main.STLineEndWidth.MED);
        tail.setLen(org.openxmlformats.schemas.drawingml.x2006.main.STLineEndLength.MED);
    }

    private static byte[] pngImage() throws IOException {
        BufferedImage image = new BufferedImage(240, 160, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g = image.createGraphics();
        try {
            g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g.setColor(new Color(0, 0, 0, 0));
            g.fillRect(0, 0, image.getWidth(), image.getHeight());
            g.setColor(new Color(0x4F, 0x81, 0xBD, 210));
            g.fillRoundRect(15, 20, 210, 120, 30, 30);
            g.setColor(new Color(0xFD, 0xE9, 0xD9));
            g.fillOval(70, 45, 100, 70);
            g.setColor(Color.WHITE);
            g.setStroke(new BasicStroke(4));
            g.drawLine(20, 20, 220, 140);
        } finally {
            g.dispose();
        }
        return imageBytes(image, "png");
    }

    private static byte[] jpegImage() throws IOException {
        BufferedImage image = new BufferedImage(240, 160, BufferedImage.TYPE_INT_RGB);
        Graphics2D g = image.createGraphics();
        try {
            g.setPaint(new GradientPaint(0, 0, new Color(0xC6, 0xE0, 0xB4),
                    240, 160, new Color(0x1F, 0x49, 0x7D)));
            g.fillRect(0, 0, image.getWidth(), image.getHeight());
            g.setColor(Color.WHITE);
            g.setFont(new java.awt.Font(java.awt.Font.SANS_SERIF, java.awt.Font.BOLD, 26));
            g.drawString("JPEG", 65, 95);
        } finally {
            g.dispose();
        }
        return imageBytes(image, "jpg");
    }

    private static byte[] imageBytes(BufferedImage image, String format) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ImageIO.write(image, format, out);
        return out.toByteArray();
    }

    private static void replaceCellGradient(Path source, Path target) throws IOException {
        String solid = "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FEDCBA\"/></patternFill></fill>";
        String gradient = "<fill><gradientFill degree=\"0\">"
                + "<stop position=\"0\"><color rgb=\"FFFFC000\"/></stop>"
                + "<stop position=\"1\"><color rgb=\"FF00B0F0\"/></stop>"
                + "</gradientFill></fill>";
        try (ZipFile zip = new ZipFile(source.toFile());
             ZipOutputStream out = new ZipOutputStream(Files.newOutputStream(target))) {
            var entries = zip.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                out.putNextEntry(new ZipEntry(entry.getName()));
                try (InputStream input = zip.getInputStream(entry)) {
                    byte[] bytes = input.readAllBytes();
                    if ("xl/styles.xml".equals(entry.getName())) {
                        String xml = new String(bytes, StandardCharsets.UTF_8);
                        if (xml.contains(solid)) {
                            xml = xml.replace(solid, gradient);
                        }
                        bytes = xml.getBytes(StandardCharsets.UTF_8);
                    }
                    out.write(bytes);
                }
                out.closeEntry();
            }
        }
    }

    private static void widths(Sheet sheet, int... chars) {
        for (int i = 0; i < chars.length; i++) {
            sheet.setColumnWidth(i, chars[i] * 256);
        }
    }

    private static void title(Sheet sheet, int row, int columns, String text, XSSFCellStyle style) {
        Row r = sheet.getRow(row);
        if (r == null) {
            r = sheet.createRow(row);
        }
        Cell cell = r.createCell(0);
        cell.setCellValue(text);
        cell.setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(row, row, 0, Math.max(0, columns - 1)));
        r.setHeightInPoints(28);
    }

    private static int section(Sheet sheet, int row, String text, XSSFCellStyle style) {
        Row r = sheet.getRow(row);
        if (r == null) {
            r = sheet.createRow(row);
        }
        Cell cell = r.createCell(0);
        cell.setCellValue(text);
        cell.setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(row, row, 0, 5));
        return row + 1;
    }

    private static int value(Sheet sheet, int row, String label, String value,
                             XSSFCellStyle valueStyle, XSSFCellStyle labelStyle) {
        Row r = sheet.getRow(row);
        if (r == null) {
            r = sheet.createRow(row);
        }
        Cell a = r.createCell(0);
        a.setCellValue(label);
        a.setCellStyle(labelStyle);
        Cell b = r.createCell(1);
        b.setCellValue(value);
        b.setCellStyle(valueStyle);
        r.setHeightInPoints(20);
        return row + 1;
    }

    private static int number(Sheet sheet, int row, String label, double value,
                              XSSFCellStyle valueStyle, XSSFCellStyle labelStyle,
                              XSSFCellStyle unused) {
        Row r = sheet.getRow(row);
        if (r == null) {
            r = sheet.createRow(row);
        }
        Cell a = r.createCell(0);
        a.setCellValue(label);
        a.setCellStyle(labelStyle);
        Cell b = r.createCell(1);
        b.setCellValue(value);
        b.setCellStyle(valueStyle);
        return row + 1;
    }

    private static XSSFFont font(XSSFWorkbook workbook, double size, boolean bold,
                                 boolean italic, boolean underline, boolean strike, int rgb) {
        return font(workbook, "Meiryo", size, bold, italic, underline, strike, rgb);
    }

    private static XSSFFont font(XSSFWorkbook workbook, String name, double size, boolean bold,
                                 boolean italic, boolean underline, boolean strike, int rgb) {
        XSSFFont font = workbook.createFont();
        font.setFontName(name);
        font.setFontHeightInPoints((short) size);
        font.setBold(bold);
        font.setItalic(italic);
        font.setUnderline(underline ? XSSFFont.U_SINGLE : XSSFFont.U_NONE);
        font.setStrikeout(strike);
        font.setColor(new XSSFColor(new byte[]{
                (byte) ((rgb >> 16) & 0xFF),
                (byte) ((rgb >> 8) & 0xFF),
                (byte) (rgb & 0xFF)}, null));
        return font;
    }

    private static int rotationExample(Sheet sheet, int row, String text, String label,
                                       int degrees, float height, XSSFCellStyle labelStyle,
                                       XSSFWorkbook workbook) {
        Row r = sheet.getRow(row);
        if (r == null) {
            r = sheet.createRow(row);
        }
        Cell value = r.createCell(0);
        value.setCellValue(text);
        value.setCellStyle(rotationStyle(workbook, degrees));
        Cell description = r.createCell(1);
        description.setCellValue(label);
        description.setCellStyle(labelStyle);
        r.setHeightInPoints(height);
        return row + 1;
    }

    private static XSSFCellStyle baseStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font(workbook, 11, false, false, false, false, 0x000000));
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private static XSSFCellStyle titleStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setFont(font(workbook, 15, true, false, false, false, 0xFFFFFF));
        style.setFillForegroundColor(new XSSFColor(new byte[]{0x1F, 0x49, 0x7D}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.LEFT);
    return style;
    }

    private static XSSFCellStyle sectionStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setFont(font(workbook, 12, true, false, false, false, 0x1F497D));
        style.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0xDD, (byte) 0xEB, (byte) 0xF7}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static XSSFCellStyle labelStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setFont(font(workbook, 11, true, false, false, false, 0x000000));
        return style;
    }

    private static XSSFCellStyle fontStyle(XSSFWorkbook workbook, double size, boolean bold,
                                           boolean italic, boolean underline, boolean strike, int rgb) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setFont(font(workbook, size, bold, italic, underline, strike, rgb));
        return style;
    }

    private static XSSFCellStyle fillStyle(XSSFWorkbook workbook, int rgb) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setFillForegroundColor(new XSSFColor(new byte[]{
                (byte) ((rgb >> 16) & 0xFF),
                (byte) ((rgb >> 8) & 0xFF),
                (byte) (rgb & 0xFF)}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private static XSSFCellStyle gradientCellStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0xFE, (byte) 0xDC, (byte) 0xBA}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static XSSFCellStyle patternStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0xC0, (byte) 0xC0, (byte) 0xC0}, null));
        style.setFillBackgroundColor(new XSSFColor(new byte[]{(byte) 0xFF, (byte) 0xFF, (byte) 0xFF}, null));
        style.setFillPattern(FillPatternType.FINE_DOTS);
        return style;
    }

    private static XSSFCellStyle borderStyle(XSSFWorkbook workbook, BorderStyle border) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setBorderTop(border);
        style.setBorderBottom(border);
        style.setBorderLeft(border);
        style.setBorderRight(border);
        return style;
    }

    private static XSSFCellStyle numberStyle(XSSFWorkbook workbook, String format) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setDataFormat(workbook.createDataFormat().getFormat(format));
        style.setAlignment(HorizontalAlignment.RIGHT);
        return style;
    }

    private static XSSFCellStyle alignmentStyle(XSSFWorkbook workbook, HorizontalAlignment horizontal,
                                                 VerticalAlignment vertical) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setAlignment(horizontal);
        style.setVerticalAlignment(vertical);
        return style;
    }

    private static XSSFCellStyle wrapStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setWrapText(true);
        return style;
    }

    private static XSSFCellStyle wrapIndentStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = wrapStyle(workbook);
        style.setIndention((short) 2);
        return style;
    }

    private static XSSFCellStyle wrapCenterStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = wrapStyle(workbook);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private static XSSFCellStyle shrinkStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setShrinkToFit(true);
        return style;
    }

    private static XSSFCellStyle rotationStyle(XSSFWorkbook workbook, int degrees) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setFont(font(workbook, "Calibri", 14, false, false, false, false, 0x000000));
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setRotation((short) degrees);
        return style;
    }

    private static XSSFCellStyle indentStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setIndention((short) 2);
        return style;
    }

    private static XSSFCellStyle hyperlinkStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = baseStyle(workbook);
        style.setFont(font(workbook, 11, false, false, true, false, 0x0563C1));
        return style;
    }

    private static XSSFCellStyle noStyle(XSSFWorkbook workbook) {
        return workbook.createCellStyle();
    }
}
