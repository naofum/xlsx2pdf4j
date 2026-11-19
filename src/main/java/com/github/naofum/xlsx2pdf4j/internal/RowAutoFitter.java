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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRow;

import java.awt.Font;
import java.awt.font.FontRenderContext;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * Expands rows that have no explicit height so wrapped cell text is visible.
 * Excel performs this auto-fit when it renders a worksheet. POI exposes the
 * default height instead, so the same measurement is done here before the page
 * layout is calculated.
 */
final class RowAutoFitter {

    private static final double CELL_HORIZONTAL_PADDING = 3.0;
    private static final double CELL_VERTICAL_PADDING = 2.0;
    // Ignore differences smaller than the rounding noise between AWT and Excel
    // font metrics. This also keeps one-line 11pt rows at the default height.
    private static final double HEIGHT_TOLERANCE = 1.5;

    private final FontResolver fonts;
    private final TextLayoutEngine textEngine;
    private final DataFormatter formatter;
    private final FontRenderContext fontContext = new FontRenderContext(null, true, true);

    RowAutoFitter(FontResolver fonts, TextLayoutEngine textEngine, DataFormatter formatter) {
        this.fonts = fonts;
        this.textEngine = textEngine;
        this.formatter = formatter;
    }

    void fit(SheetLayout layout) {
        Set<Long> mergedCells = mergedCells(layout);
        double[] heights = layout.rowHeights.clone();
        for (int rowIndex = layout.firstRow; rowIndex <= layout.lastRow; rowIndex++) {
            XSSFRow row = layout.sheet.getRow(rowIndex);
            if (row == null || row.getZeroHeight() || hasExplicitHeight(row)) {
                continue;
            }
            double required = requiredHeight(layout, row, rowIndex, mergedCells);
            if (required > heights[rowIndex] + HEIGHT_TOLERANCE) {
                heights[rowIndex] = required;
            }
        }
        layout.applyRowHeights(heights);
    }

    private double requiredHeight(SheetLayout layout, XSSFRow row, int rowIndex,
                                  Set<Long> mergedCells) {
        double required = 0;
        for (int column = layout.firstColumn; column <= layout.lastColumn; column++) {
            if (column < 0 || column >= layout.columnWidths.length
                    || layout.columnWidths[column] <= 0
                    || layout.sheet.isColumnHidden(column)
                    || mergedCells.contains(key(rowIndex, column))) {
                continue;
            }
            Cell cell = row.getCell(column);
            if (cell == null || cell.getCellStyle() == null) {
                continue;
            }
            XSSFCellStyle style = cell.getCellStyle() instanceof XSSFCellStyle xssf ? xssf : null;
            if (style == null || !style.getWrapText() || style.getShrinkToFit()
                    || style.getRotation() != 0) {
                continue;
            }
            String text = displayText(cell);
            if (text == null || text.isBlank()) {
                continue;
            }
            XSSFFont excelFont = style.getFont();
            String fontName = excelFont == null ? "MS PGothic" : excelFont.getFontName();
            double fontSize = excelFont == null ? 11 : excelFont.getFontHeightInPoints();
            boolean bold = excelFont != null && excelFont.getBold();
            boolean italic = excelFont != null && excelFont.getItalic();
            Font font = fonts.resolveForText(fontName, fontSize, bold, italic, text);
            double indent = style.getIndention() * fonts.scaleSize(fontSize) * 0.6;
            double width = Math.max(1, layout.columnWidths[column]
                    - CELL_HORIZONTAL_PADDING - indent);
            List<String> lines = textEngine.wrapLines(text, font, fontContext, width, true);
            double lineHeight = font.getLineMetrics("Ag", fontContext).getHeight();
            if (lineHeight <= 0) {
                lineHeight = Math.max(1, font.getSize2D());
            }
            double contentHeight = Math.max(1, lines.size()) * lineHeight;
            double height = SheetLayout.storedRowHeight(contentHeight + CELL_VERTICAL_PADDING);
            required = Math.max(required, height);
        }
        return required;
    }

    private String displayText(Cell cell) {
        CellType type = cell.getCellType();
        if (type == CellType.STRING) {
            return formatter.formatCellValue(cell);
        }
        if (type == CellType.FORMULA && !hasUncachedFormula(cell)
                && cell.getCachedFormulaResultType() == CellType.STRING) {
            return formatter.formatCellValue(cell);
        }
        return null;
    }

    private static boolean hasUncachedFormula(Cell cell) {
        return cell instanceof XSSFCell xssf
                && xssf.getCellType() == CellType.FORMULA
                && xssf.getCTCell().isSetF()
                && !xssf.getCTCell().isSetV();
    }

    private static boolean hasExplicitHeight(XSSFRow row) {
        CTRow ct = row.getCTRow();
        if (ct == null) {
            return false;
        }
        if (ct.isSetHt()) {
            return true;
        }
        return ct.isSetCustomHeight() && ct.getCustomHeight();
    }

    private static Set<Long> mergedCells(SheetLayout layout) {
        Set<Long> cells = new HashSet<>();
        for (CellRangeAddress region : layout.mergedRegions) {
            int firstRow = Math.max(layout.firstRow, region.getFirstRow());
            int lastRow = Math.min(layout.lastRow, region.getLastRow());
            int firstColumn = Math.max(layout.firstColumn, region.getFirstColumn());
            int lastColumn = Math.min(layout.lastColumn, region.getLastColumn());
            for (int row = firstRow; row <= lastRow; row++) {
                for (int column = firstColumn; column <= lastColumn; column++) {
                    cells.add(key(row, column));
                }
            }
        }
        return cells;
    }

    private static long key(int row, int column) {
        return ((long) row << 32) | (column & 0xffffffffL);
    }
}
