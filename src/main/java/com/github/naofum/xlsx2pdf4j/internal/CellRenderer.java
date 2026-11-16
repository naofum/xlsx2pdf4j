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

package com.github.naofum.xlsx2pdf4j.internal;

import com.github.naofum.xlsx2pdf4j.api.ConversionWarning;
import com.github.naofum.xlsx2pdf4j.api.WarningCode;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.Paint;
import java.awt.Shape;
import java.awt.Stroke;
import java.awt.geom.Rectangle2D;
import java.time.LocalDateTime;
import java.time.chrono.JapaneseChronology;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

/**
 * Draws worksheet cells. Coordinates and widths are in points; the page
 * renderer applies the DPI transform.
 */
public final class CellRenderer {

    private final XSSFWorkbook workbook;
    private final XSSFSheet sheet;
    private final SheetLayout layout;
    private final FontResolver fonts;
    private final TextLayoutEngine textEngine;
    private final DataFormatter formatter;
    private final FormulaEvaluator evaluator;
    private final List<ConversionWarning> warnings;
    private final Set<String> warnedFormulaCells = new HashSet<>();
    private final Map<Integer, CellGradient> gradients = new HashMap<>();
    private final Map<Long, CellRangeAddress> mergedByCell = new HashMap<>();

    public CellRenderer(XSSFWorkbook workbook,
                        SheetLayout layout,
                        FontResolver fonts,
                        TextLayoutEngine textEngine,
                        DataFormatter formatter,
                        FormulaEvaluator evaluator,
                        List<ConversionWarning> warnings) {
        this.workbook = workbook;
        this.sheet = layout.sheet;
        this.layout = layout;
        this.fonts = fonts;
        this.textEngine = textEngine;
        this.formatter = formatter;
        this.evaluator = evaluator;
        this.warnings = warnings;
        for (CellRangeAddress region : sheet.getMergedRegions()) {
            for (int r = region.getFirstRow(); r <= region.getLastRow(); r++) {
                for (int c = region.getFirstColumn(); c <= region.getLastColumn(); c++) {
                    mergedByCell.put(key(r, c), region);
                }
            }
        }
    }

    public void drawRange(Graphics2D g2,
                          int firstRow,
                          int lastRow,
                          int firstColumn,
                          int lastColumn,
                          double baseX,
                          double baseY,
                          double scale,
                          java.util.List<PlacedText> textLayer) {
        Shape oldClip = g2.getClip();
        double clipX = baseX;
        double clipY = baseY;
        double clipW = (layout.columnX[lastColumn + 1] - layout.columnX[firstColumn]) * scale;
        double clipH = (layout.renderRowY[lastRow + 1] - layout.renderRowY[firstRow]) * scale;
        // Border strokes are centered on the cell edge. A little padding keeps
        // the line visible when the cell sits exactly on a page boundary.
        double padding = 2.0;
        g2.clip(new Rectangle2D.Double(clipX - padding, clipY - padding,
                Math.max(0, clipW) + padding * 2, Math.max(0, clipH) + padding * 2));
        try {
            // Draw all fills first, then all borders, then text. Drawing a
            // cell's fill and border together lets the next row/column's fill
            // cover the shared border (a white fill makes the border vanish).
            for (CellPass pass : CellPass.values()) {
                for (int row = firstRow; row <= lastRow; row++) {
                    if (row < 0 || row >= layout.rowHeights.length || layout.rowHeights[row] <= 0) {
                        continue;
                    }
                    Row sheetRow = sheet.getRow(row);
                    for (int col = firstColumn; col <= lastColumn; col++) {
                        if (col < 0 || col >= layout.columnWidths.length || layout.columnWidths[col] <= 0) {
                            continue;
                        }
                        CellRangeAddress merged = mergedByCell.get(key(row, col));
                        if (merged != null) {
                            if (merged.getFirstRow() != row || merged.getFirstColumn() != col) {
                                continue;
                            }
                            int endRow = Math.min(merged.getLastRow(), lastRow);
                            int endCol = Math.min(merged.getLastColumn(), lastColumn);
                            drawCell(g2, sheetRow, row, col, endRow, endCol, baseX, baseY, scale,
                                    firstRow, firstColumn, lastColumn, true, pass, textLayer);
                        } else {
                            drawCell(g2, sheetRow, row, col, row, col, baseX, baseY, scale,
                                    firstRow, firstColumn, lastColumn, false, pass, textLayer);
                        }
                    }
                }
            }
        } finally {
            g2.setClip(oldClip);
        }
    }

    private void drawCell(Graphics2D g2,
                          Row sheetRow,
                          int row,
                          int col,
                          int endRow,
                          int endCol,
                          double baseX,
                          double baseY,
                          double scale,
                          int pageFirstRow,
                          int pageFirstColumn,
                          int pageLastColumn,
                          boolean mergedCell,
                          CellPass pass,
                          java.util.List<PlacedText> textLayer) {
        Cell cell = sheetRow == null ? null : sheetRow.getCell(col);
        CellStyle baseStyle = cell == null ? null : cell.getCellStyle();
        XSSFCellStyle style = baseStyle instanceof XSSFCellStyle s ? s : null;

        double x = baseX + (layout.columnX[col] - layout.columnX[pageFirstColumn]) * scale;
        double y = baseY + (layout.renderRowY[row] - layout.renderRowY[pageFirstRow]) * scale;
        double width = (layout.columnX[endCol + 1] - layout.columnX[col]) * scale;
        double height = (layout.renderRowY[endRow + 1] - layout.renderRowY[row]) * scale;

        if (pass == CellPass.FILL) {
            if (style != null) {
                drawFill(g2, style, x, y, width, height);
            }
            return;
        }
        if (pass == CellPass.BORDER) {
            if (style == null) {
                drawIncomingTopBorder(g2, null, row, col, pageFirstRow, x, y, width, scale);
            } else if (mergedCell) {
                drawMergedBorders(g2, style, row, col, endRow, endCol, x, y, width, height,
                        pageFirstRow, pageFirstColumn, scale);
            } else {
                drawBorders(g2, style, row, col, pageFirstRow, pageFirstColumn, pageLastColumn,
                    x, y, width, height, scale);
            }
            return;
        }
        if (style == null) {
            return;
        }

        String value = formatCellValue(cell);
        if (value == null || value.isEmpty()) {
            return;
        }

        XSSFFont excelFont = style.getFont();
        boolean bold = excelFont != null && excelFont.getBold();
        boolean italic = excelFont != null && excelFont.getItalic();
        boolean underline = excelFont != null && excelFont.getUnderline() != 0;
        boolean strike = excelFont != null && excelFont.getStrikeout();
        double fontSize = excelFont == null ? 11 : excelFont.getFontHeightInPoints();
        String fontName = excelFont == null ? "MS PGothic" : excelFont.getFontName();
        Color fontColor = excelFont == null ? Color.BLACK : ColorUtil.toAwt(excelFont.getXSSFColor(), workbook, Color.BLACK);
        Font font = fonts.resolveForText(fontName, fontSize * scale, bold, italic, value);

        double renderSize = fonts.scaleSize(fontSize) * scale;
        if (textLayer != null) {
            // Always use the full cell value here, not the clipped/wrapped
            // visual fragment. The layer is invisible and should remain
            // searchable even when the raster output clips the text.
            double baseline = y + Math.max(1, Math.min(height - 1, renderSize + 1));
            textLayer.add(new PlacedText(value, x + 1.5, baseline, Math.max(1, renderSize)));
        }

        HorizontalAlignment horizontal = effectiveHorizontal(cell, style.getAlignment());
        VerticalAlignment vertical = style.getVerticalAlignment();
        if (vertical == null) {
            vertical = VerticalAlignment.BOTTOM;
        }
        double indent = style.getIndention() * renderSize * 0.6;
        Rectangle2D textRect = new Rectangle2D.Double(
                x + 1.5,
                y + 0.75,
                Math.max(1, width - 3),
                Math.max(1, height - 1.5));

        int overflowLastCol = endCol;
        if (!mergedCell && horizontal == HorizontalAlignment.CENTER_SELECTION) {
            // Excel starts a center-across-selection span at the value cell and
            // extends it over following blank cells. It must not absorb the next
            // value cell, otherwise every label on the row centers over the same
            // full-width range.
            int rangeLast = endCol;
            while (rangeLast < pageLastColumn && isCenterContinuous(row, rangeLast + 1)
                    && isEmptyCell(row, rangeLast + 1)) {
                rangeLast++;
            }
            if (rangeLast > endCol) {
                double rangeWidth = (layout.columnX[rangeLast + 1] - layout.columnX[col]) * scale;
                textRect = new Rectangle2D.Double(x + 1.5, y + 0.75,
                        Math.max(1, rangeWidth - 3), Math.max(1, height - 1.5));
            }
            overflowLastCol = rangeLast;
            horizontal = HorizontalAlignment.CENTER;
        }

        if (!mergedCell && !style.getWrapText() && !style.getShrinkToFit() && isTextCell(cell)) {
            if (horizontal == HorizontalAlignment.CENTER) {
                // Excel also overflows plain centered text symmetrically into
                // blank cells, not only center-across-selection spans.
                textRect = expandCenteredForOverflow(row, col, overflowLastCol,
                        pageFirstColumn, pageLastColumn, g2, value, font, textRect, scale);
            } else if (!isCentered(horizontal)) {
                textRect = expandForOverflow(cell, row, col, endCol, pageLastColumn, g2, value, font,
                        textRect, horizontal, scale);
            }
        }

        java.util.List<TextLayoutEngine.StyledRun> richRuns =
                richRuns(cell, excelFont, font, fontColor, scale);
        if (richRuns != null) {
            textEngine.drawRichText(g2, richRuns, textRect, horizontal, vertical,
                    style.getWrapText(), style.getRotation(), indent);
            return;
        }

        textEngine.drawText(g2, value, font, fontColor, textRect, horizontal, vertical,
                style.getWrapText(), style.getShrinkToFit(), style.getRotation(), indent, underline, strike);
    }

    private java.util.List<TextLayoutEngine.StyledRun> richRuns(Cell cell, XSSFFont baseFont,
                                                               Font baseAwtFont, Color baseColor,
                                                               double layoutScale) {
        if (cell == null || cell.getCellType() != org.apache.poi.ss.usermodel.CellType.STRING) {
            return null;
        }
        XSSFRichTextString rich;
        try {
            org.apache.poi.ss.usermodel.RichTextString value = cell.getRichStringCellValue();
            if (!(value instanceof XSSFRichTextString xssf)) {
                return null;
            }
            rich = xssf;
        } catch (RuntimeException e) {
            return null;
        }
        if (rich == null || rich.numFormattingRuns() == 0) {
            return null;
        }
        java.util.List<TextLayoutEngine.StyledRun> runs = new java.util.ArrayList<>();
        int position = 0;
        int count = rich.numFormattingRuns();
        for (int i = 0; i < count; i++) {
            int start = rich.getIndexOfFormattingRun(i);
            if (start > position) {
                runs.add(new TextLayoutEngine.StyledRun(rich.getString().substring(position, start),
                        baseAwtFont, baseColor,
                        baseFont != null && baseFont.getUnderline() != 0,
                        baseFont != null && baseFont.getStrikeout()));
            }
            int end = i + 1 < count ? rich.getIndexOfFormattingRun(i + 1) : rich.length();
            if (end > start) {
                XSSFFont runFont = rich.getFontOfFormattingRun(i);
                XSSFFont effective = runFont == null ? baseFont : runFont;
                boolean bold = effective != null && effective.getBold();
                boolean italic = effective != null && effective.getItalic();
                boolean underline = effective != null && effective.getUnderline() != 0;
                boolean strike = effective != null && effective.getStrikeout();
                double size = effective == null ? 11 : effective.getFontHeightInPoints();
                String name = effective == null ? "MS PGothic" : effective.getFontName();
                Color color = effective == null || effective.getXSSFColor() == null
                        ? baseColor
                        : ColorUtil.toAwt(effective.getXSSFColor(), workbook, baseColor);
                String runText = rich.getString().substring(start, end);
                runs.add(new TextLayoutEngine.StyledRun(runText,
                        fonts.resolveForText(name, size * layoutScale, bold, italic, runText),
                        color, underline, strike));
            }
            position = Math.max(position, end);
        }
        if (position < rich.length()) {
            runs.add(new TextLayoutEngine.StyledRun(rich.getString().substring(position),
                    baseAwtFont, baseColor,
                    baseFont != null && baseFont.getUnderline() != 0,
                    baseFont != null && baseFont.getStrikeout()));
        }
        return runs.isEmpty() ? null : runs;
    }

    private HorizontalAlignment effectiveHorizontal(Cell cell, HorizontalAlignment styleAlignment) {
        if (cell == null || styleAlignment == null || styleAlignment == HorizontalAlignment.GENERAL) {
            if (cell == null) {
                return HorizontalAlignment.LEFT;
            }
            if (cell.getCellType() == CellType.FORMULA) {
                return switch (cell.getCachedFormulaResultType()) {
                    case NUMERIC -> HorizontalAlignment.RIGHT;
                    case BOOLEAN, ERROR -> HorizontalAlignment.CENTER;
                    default -> HorizontalAlignment.LEFT;
                };
            }
            return switch (cell.getCellType()) {
                case NUMERIC -> HorizontalAlignment.RIGHT;
                case BOOLEAN, ERROR -> HorizontalAlignment.CENTER;
                default -> HorizontalAlignment.LEFT;
            };
        }
        return styleAlignment;
    }

    private boolean isTextCell(Cell cell) {
        if (cell == null) {
            return true;
        }
        return switch (cell.getCellType()) {
            case STRING -> true;
            case FORMULA -> cell.getCachedFormulaResultType() == CellType.STRING;
            default -> false;
        };
    }

    private boolean isCenterContinuous(int row, int col) {
        Row sheetRow = sheet.getRow(row);
        Cell cell = sheetRow == null ? null : sheetRow.getCell(col);
        if (cell == null || !(cell.getCellStyle() instanceof XSSFCellStyle style)) {
            return false;
        }
        return style.getAlignment() == HorizontalAlignment.CENTER_SELECTION;
    }

    private boolean isCentered(HorizontalAlignment alignment) {
        return alignment == HorizontalAlignment.CENTER
                || alignment == HorizontalAlignment.CENTER_SELECTION
                || alignment == HorizontalAlignment.JUSTIFY
                || alignment == HorizontalAlignment.DISTRIBUTED
                || alignment == HorizontalAlignment.FILL;
    }

    private Rectangle2D expandForOverflow(Cell cell, int row, int col, int endCol, int pageLastColumn,
                                          Graphics2D g2, String value, Font font,
                                          Rectangle2D rect, HorizontalAlignment alignment, double scale) {
        double needed = textEngine.stringWidth(g2, value, font) - rect.getWidth();
        if (needed <= 0) {
            return rect;
        }
        double x = rect.getX();
        double y = rect.getY();
        double width = rect.getWidth();
        double height = rect.getHeight();

        if (alignment == HorizontalAlignment.RIGHT) {
            int c = col - 1;
            while (c >= 0 && needed > 0 && isEmptyForOverflow(row, c)) {
                double add = layout.columnWidths[c] * scale;
                x -= add;
                width += add;
                needed -= add;
                c--;
            }
        } else {
            int c = endCol + 1;
            while (c <= pageLastColumn && needed > 0 && isEmptyForOverflow(row, c)) {
                double add = layout.columnWidths[c] * scale;
                width += add;
                needed -= add;
                c++;
            }
        }
        return new Rectangle2D.Double(x, y, width, height);
    }

    /**
     * Center-across-selection overflow is symmetric: Excel lets the text spill
     * into blank cells on both sides while keeping the span centered.
     */
    private Rectangle2D expandCenteredForOverflow(int row, int firstCol, int lastCol,
                                                  int pageFirstColumn, int pageLastColumn,
                                                  Graphics2D g2, String value, Font font,
                                                  Rectangle2D rect, double scale) {
        double needed = textEngine.stringWidth(g2, value, font) - rect.getWidth();
        if (needed <= 0) {
            return rect;
        }
        double leftAvailable = 0;
        for (int c = firstCol - 1; c >= pageFirstColumn && isEmptyForOverflow(row, c); c--) {
            leftAvailable += layout.columnWidths[c] * scale;
        }
        double rightAvailable = 0;
        for (int c = lastCol + 1; c <= pageLastColumn && isEmptyForOverflow(row, c); c++) {
            rightAvailable += layout.columnWidths[c] * scale;
        }
        // Excel keeps the text centered on the cell and lets it spill equally
        // into blank cells on both sides. Do not shift the centre when one
        // side is shorter; clip that side instead.
        double half = needed / 2d;
        double left = Math.min(half, leftAvailable);
        double right = Math.min(half, rightAvailable);
        return new Rectangle2D.Double(rect.getX() - left, rect.getY(),
                rect.getWidth() + left + right, rect.getHeight());
    }

    /**
     * Returns the value cell and the following blank center-across-selection
     * cells that Excel treats as one display span. Returns null for cells that
     * do not belong to such a span.
     */
    private int[] centerContinuousGroup(int row, int col, int pageLastColumn) {
        if (!isCenterContinuous(row, col)) {
            return null;
        }
        int start = col;
        if (isEmptyCell(row, col)) {
            int c = col - 1;
            while (c >= 0 && isCenterContinuous(row, c)) {
                if (!isEmptyCell(row, c)) {
                    start = c;
                    break;
                }
                c--;
            }
            if (start == col) {
                return null;
            }
        }
        int end = start;
        while (end < pageLastColumn && isCenterContinuous(row, end + 1)
                && isEmptyCell(row, end + 1)) {
            end++;
        }
        if (end <= start) {
            return null;
        }
        return new int[]{start, end};
    }

    private boolean isEmptyCell(int row, int col) {
        Row sheetRow = sheet.getRow(row);
        Cell cell = sheetRow == null ? null : sheetRow.getCell(col);
        if (cell == null) {
            return true;
        }
        String value = formatCellValue(cell);
        return value == null || value.isEmpty();
    }

    private boolean isEmptyForOverflow(int row, int col) {
        return isEmptyCell(row, col);
    }

    private String formatCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        if (isJapaneseBuiltinDateFormat(cell)) {
            try {
                LocalDateTime value = DateUtil.getLocalDateTime(
                        cell.getNumericCellValue(), workbook.isDate1904());
                String formatted = formatJapaneseBuiltinDate(value, cell.getCellStyle().getDataFormat());
                if (formatted != null) {
                    return formatted;
                }
            } catch (RuntimeException e) {
                // Fall through to DataFormatter for unusual/invalid values.
            }
        }
        if (hasUncachedFormula(cell)) {
            try {
                return formatter.formatCellValue(cell, evaluator);
            } catch (RuntimeException e) {
                warnUncachedFormula(cell, e);
                return "#N/A";
            }
        }
        return formatter.formatCellValue(cell, evaluator);
    }

    private static boolean hasUncachedFormula(Cell cell) {
        return cell instanceof XSSFCell xssf
                && xssf.getCellType() == CellType.FORMULA
                && xssf.getCTCell().isSetF()
                && !xssf.getCTCell().isSetV();
    }

    private void warnUncachedFormula(Cell cell, RuntimeException failure) {
        String address = cell.getAddress().formatAsString();
        if (!warnedFormulaCells.add(address)) {
            return;
        }
        String detail = failure.getMessage() == null ? failure.getClass().getSimpleName()
                : failure.getMessage();
        warnings.add(new ConversionWarning(WarningCode.FORMULA_NO_CACHE,
                sheet.getSheetName(), address,
                "formula result is not cached and could not be evaluated: " + detail));
    }

    private static boolean isJapaneseBuiltinDateFormat(Cell cell) {
        short format = cell.getCellStyle().getDataFormat();
        if (format < 27 || format > 36) {
            return false;
        }
        CellType type = cell.getCellType();
        return type == CellType.NUMERIC
                || (type == CellType.FORMULA
                    && cell.getCachedFormulaResultType() == CellType.NUMERIC);
    }

    /**
     * POI exposes Excel's Japanese built-in date/time formats (27-36) as
     * reserved-0xNN. Format them with the same meaning Excel uses.
     */
    private static String formatJapaneseBuiltinDate(LocalDateTime value, int format) {
        DateTimeFormatter formatter;
        boolean japaneseEra = false;
        switch (format) {
            case 27 -> {
                formatter = DateTimeFormatter.ofPattern("y/M/d", Locale.JAPAN);
                japaneseEra = true;
            }
            case 28, 29 -> {
                formatter = DateTimeFormatter.ofPattern("y年M月d日", Locale.JAPAN);
                japaneseEra = true;
            }
            case 30 -> formatter = DateTimeFormatter.ofPattern("M/d/yy", Locale.JAPAN);
            case 31 -> formatter = DateTimeFormatter.ofPattern("yyyy年M月d日", Locale.JAPAN);
            case 32 -> formatter = DateTimeFormatter.ofPattern("H時mm分", Locale.JAPAN);
            case 33 -> formatter = DateTimeFormatter.ofPattern("H時mm分ss秒", Locale.JAPAN);
            case 34 -> formatter = DateTimeFormatter.ofPattern("yyyy年M月", Locale.JAPAN);
            case 35 -> formatter = DateTimeFormatter.ofPattern("M月d日", Locale.JAPAN);
            case 36 -> {
                formatter = DateTimeFormatter.ofPattern("y年M月", Locale.JAPAN);
                japaneseEra = true;
            }
            default -> {
                return null;
            }
        }
        if (japaneseEra) {
            formatter = formatter.withChronology(JapaneseChronology.INSTANCE);
        }
        return formatter.format(value);
    }

    private void drawFill(Graphics2D g2, XSSFCellStyle style, double x, double y, double width, double height) {
        CellGradient gradient = gradient(style);
        if (gradient != null) {
            Paint paint = gradient.paint(new Rectangle2D.Double(x, y, width, height));
            if (paint != null) {
                fillRect(g2, paint, x, y, width, height);
                return;
            }
        }
        FillPatternType pattern = style.getFillPattern();
        if (pattern == null || pattern == FillPatternType.NO_FILL) {
            return;
        }
        Color color = ColorUtil.toAwt(style.getFillForegroundXSSFColor(), workbook, null);
        if (color == null || color.getAlpha() == 0) {
            return;
        }
        fillRect(g2, color, x, y, width, height);
    }

    private CellGradient gradient(XSSFCellStyle style) {
        if (style == null || style.getCoreXf() == null) {
            return null;
        }
        int fillId = (int) style.getCoreXf().getFillId();
        if (gradients.containsKey(fillId)) {
            return gradients.get(fillId);
        }
        CellGradient gradient = CellGradient.read(workbook, style);
        gradients.put(fillId, gradient);
        return gradient;
    }

    private static void fillRect(Graphics2D g2, Paint paint,
                                 double x, double y, double width, double height) {
        g2.setPaint(paint);
        // Adjacent cell fills share an edge. With antialiasing on, that edge is
        // blended twice and appears as a thin darker line (especially visible
        // in a continuous colored band). Cell fills are plain rectangles, so
        // render them without antialiasing to keep the band seamless.
        Object old = g2.getRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING);
        g2.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING,
                java.awt.RenderingHints.VALUE_ANTIALIAS_OFF);
        try {
            g2.fill(new Rectangle2D.Double(x, y, width, height));
        } finally {
            g2.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING,
                    old == null ? java.awt.RenderingHints.VALUE_ANTIALIAS_ON : old);
        }
    }

    private void drawMergedBorders(Graphics2D g2, XSSFCellStyle style,
                                   int firstRow, int firstColumn, int lastRow, int lastColumn,
                                   double x, double y, double width, double height,
                                   int pageFirstRow, int pageFirstColumn, double scale) {
        // The origin cell is a useful fallback, but Excel can take the outer
        // border from any cell along the merged edge. Draw both the origin's
        // outer border and the edge cells' borders to stay close.
        drawBorder(g2, style.getBorderTop(), borderColor(style, XSSFCellBorder.BorderSide.TOP),
                x, y, x + width, y, scale);
        drawBorder(g2, style.getBorderBottom(), borderColor(style, XSSFCellBorder.BorderSide.BOTTOM),
                x, y + height, x + width, y + height, scale);
        drawBorder(g2, style.getBorderLeft(), borderColor(style, XSSFCellBorder.BorderSide.LEFT),
                x, y, x, y + height, scale);
        drawBorder(g2, style.getBorderRight(), borderColor(style, XSSFCellBorder.BorderSide.RIGHT),
                x + width, y, x + width, y + height, scale);

        for (int c = firstColumn; c <= lastColumn; c++) {
            XSSFCellStyle edge = styleAt(firstRow, c, style);
            double x1 = x + (layout.columnX[c] - layout.columnX[firstColumn]) * scale;
            double x2 = x + (layout.columnX[c + 1] - layout.columnX[firstColumn]) * scale;
            drawBorder(g2, edge.getBorderTop(), borderColor(edge, XSSFCellBorder.BorderSide.TOP),
                    x1, y, x2, y, scale);
            XSSFCellStyle bottom = styleAt(lastRow, c, style);
            drawBorder(g2, bottom.getBorderBottom(), borderColor(bottom, XSSFCellBorder.BorderSide.BOTTOM),
                    x1, y + height, x2, y + height, scale);
        }
        for (int r = firstRow; r <= lastRow; r++) {
            XSSFCellStyle edge = styleAt(r, firstColumn, style);
            double y1 = y + (layout.renderRowY[r] - layout.renderRowY[firstRow]) * scale;
            double y2 = y + (layout.renderRowY[r + 1] - layout.renderRowY[firstRow]) * scale;
            drawBorder(g2, edge.getBorderLeft(), borderColor(edge, XSSFCellBorder.BorderSide.LEFT),
                    x, y1, x, y2, scale);
            XSSFCellStyle right = styleAt(r, lastColumn, style);
            drawBorder(g2, right.getBorderRight(), borderColor(right, XSSFCellBorder.BorderSide.RIGHT),
                    x + width, y1, x + width, y2, scale);
        }
    }

    private XSSFCellStyle styleAt(int row, int column, XSSFCellStyle fallback) {
        Row sheetRow = sheet.getRow(row);
        Cell cell = sheetRow == null ? null : sheetRow.getCell(column);
        if (cell != null && cell.getCellStyle() instanceof XSSFCellStyle style) {
            return style;
        }
        return fallback;
    }

    private void drawBorders(Graphics2D g2, XSSFCellStyle style, int row, int col,
                             int pageFirstRow, int pageFirstColumn, int pageLastColumn,
                             double x, double y, double width, double height, double scale) {
        int[] group = centerContinuousGroup(row, col, pageLastColumn);
        int groupStart = group == null ? col : Math.max(group[0], pageFirstColumn);
        int groupEnd = group == null ? col : Math.min(group[1], pageLastColumn);

        drawBorder(g2, style.getBorderTop(), borderColor(style, XSSFCellBorder.BorderSide.TOP), x, y, x + width, y, scale);
        drawBorder(g2, style.getBorderBottom(), borderColor(style, XSSFCellBorder.BorderSide.BOTTOM), x, y + height, x + width, y + height, scale);
        drawIncomingTopBorder(g2, style, row, col, pageFirstRow, x, y, width, scale);
        // Excel suppresses the internal vertical borders of a center-across
        // span and draws only the span's outer edges.
        if (col <= groupStart) {
            drawBorder(g2, style.getBorderLeft(), borderColor(style, XSSFCellBorder.BorderSide.LEFT), x, y, x, y + height, scale);
        }
        if (col >= groupEnd) {
            drawBorder(g2, style.getBorderRight(), borderColor(style, XSSFCellBorder.BorderSide.RIGHT), x + width, y, x + width, y + height, scale);
        }
    }

    /**
     * Excel prints the bottom border of the row above at the top edge of the
     * next page. Without this, a row that relies on the previous row's bottom
     * border (top border is NONE) loses that line at a page break.
     */
    private void drawIncomingTopBorder(Graphics2D g2, XSSFCellStyle style, int row, int col,
                                       int pageFirstRow, double x, double y,
                                       double width, double scale) {
        if (row != pageFirstRow || row <= 0
                || row >= layout.rowHeights.length || layout.rowHeights[row - 1] <= 0) {
            return;
        }
        BorderStyle top = style == null ? null : style.getBorderTop();
        if (top != null && top != BorderStyle.NONE) {
            return;
        }
        XSSFCellStyle above = styleAt(row - 1, col, null);
        if (above == null) {
            return;
        }
        drawBorder(g2, above.getBorderBottom(),
                borderColor(above, XSSFCellBorder.BorderSide.BOTTOM),
                x, y, x + width, y, scale);
    }

    private Color borderColor(XSSFCellStyle style, XSSFCellBorder.BorderSide side) {
        try {
            return ColorUtil.toAwt(style.getBorderColor(side), workbook, Color.BLACK);
        } catch (RuntimeException e) {
            return Color.BLACK;
        }
    }

    private void drawBorder(Graphics2D g2, BorderStyle style, Color color,
                            double x1, double y1, double x2, double y2, double scale) {
        if (style == null || style == BorderStyle.NONE) {
            return;
        }
        float width = switch (style) {
            case HAIR -> 0.25f;
            case MEDIUM -> 1.0f;
            case THICK -> 1.5f;
            case DOUBLE -> 0.5f;
            default -> 0.5f;
        };
        g2.setColor(color == null ? Color.BLACK : color);
        if (style == BorderStyle.DOUBLE) {
            Stroke old = g2.getStroke();
            g2.setStroke(new BasicStroke(width));
            g2.draw(new java.awt.geom.Line2D.Double(x1, y1, x2, y2));
            double offset = 1.5;
            if (Math.abs(y1 - y2) < 0.01) {
                g2.draw(new java.awt.geom.Line2D.Double(x1, y1 + offset, x2, y2 + offset));
            } else {
                g2.draw(new java.awt.geom.Line2D.Double(x1 + offset, y1, x2 + offset, y2));
            }
            g2.setStroke(old);
            return;
        }
        Stroke old = g2.getStroke();
        if (style == BorderStyle.DASHED || style == BorderStyle.DASH_DOT || style == BorderStyle.DASH_DOT_DOT) {
            g2.setStroke(new BasicStroke(width, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER,
                    10f, new float[]{3f, 2f}, 0f));
        } else if (style == BorderStyle.DOTTED) {
            g2.setStroke(new BasicStroke(width, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER,
                    10f, new float[]{1f, 2f}, 0f));
        } else {
            g2.setStroke(new BasicStroke(width));
        }
        g2.draw(new java.awt.geom.Line2D.Double(x1, y1, x2, y2));
        g2.setStroke(old);
    }

    private enum CellPass {
        FILL,
        BORDER,
        TEXT
    }

    private static long key(int row, int col) {
        return ((long) row << 20) | (col & 0xFFFFFL);
    }
}
