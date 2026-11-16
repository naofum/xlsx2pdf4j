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

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Geometry and print settings for one worksheet.
 */
public final class SheetLayout {

    // Excel's PDF export renders row heights at 17/18 of the stored height
    // (for example, a 15pt row occupies 14.17pt on the page). Column widths
    // are not affected by this factor.
    private static final double ROW_HEIGHT_SCALE = 17d / 18d;

    public final XSSFWorkbook workbook;
    public final XSSFSheet sheet;
    public final int sheetIndex;
    public final String name;
    public final int firstRow;
    public final int lastRow;
    public final int firstColumn;
    public final int lastColumn;
    public final int contentFirstRow;
    public final int contentFirstColumn;
    public final int titleFirstRow;
    public final int titleLastRow;
    public final int titleFirstColumn;
    public final int titleLastColumn;
    public final double[] columnWidths;
    public final double[] columnX;
    public final double[] rowHeights;
    public final double[] rowY;
    public final double[] renderRowY;
    public final PrintSettingsModel print;
    public final List<CellRangeAddress> mergedRegions;

    private SheetLayout(XSSFWorkbook workbook,
                        XSSFSheet sheet,
                        int sheetIndex,
                        int firstRow,
                        int lastRow,
                        int firstColumn,
                        int lastColumn,
                        int contentFirstRow,
                        int contentFirstColumn,
                        int titleFirstRow,
                        int titleLastRow,
                        int titleFirstColumn,
                        int titleLastColumn,
                        double[] columnWidths,
                        double[] columnX,
                        double[] rowHeights,
                        double[] rowY,
                        double[] renderRowY,
                        PrintSettingsModel print,
                        List<CellRangeAddress> mergedRegions) {
        this.workbook = workbook;
        this.sheet = sheet;
        this.sheetIndex = sheetIndex;
        this.name = sheet.getSheetName();
        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstColumn = firstColumn;
        this.lastColumn = lastColumn;
        this.contentFirstRow = contentFirstRow;
        this.contentFirstColumn = contentFirstColumn;
        this.titleFirstRow = titleFirstRow;
        this.titleLastRow = titleLastRow;
        this.titleFirstColumn = titleFirstColumn;
        this.titleLastColumn = titleLastColumn;
        this.columnWidths = columnWidths;
        this.columnX = columnX;
        this.rowHeights = rowHeights;
        this.rowY = rowY;
        this.renderRowY = renderRowY;
        this.print = print;
        this.mergedRegions = List.copyOf(mergedRegions);
    }

    public static SheetLayout read(XSSFWorkbook workbook, int sheetIndex, XSSFSheet sheet,
                                   double columnWidthScale) {
        int[] area = readPrintArea(workbook, sheetIndex, sheet);
        int firstRow = area[0];
        int lastRow = area[1];
        int firstColumn = area[2];
        int lastColumn = area[3];

        CellRangeAddress titleRows = sheet.getRepeatingRows();
        CellRangeAddress titleColumns = sheet.getRepeatingColumns();
        int titleFirstRow = titleRows == null ? -1 : titleRows.getFirstRow();
        int titleLastRow = titleRows == null ? -1 : titleRows.getLastRow();
        int titleFirstColumn = titleColumns == null ? -1 : titleColumns.getFirstColumn();
        int titleLastColumn = titleColumns == null ? -1 : titleColumns.getLastColumn();

        int contentFirstRow = firstRow;
        if (titleRows != null && titleRows.getFirstRow() <= firstRow) {
            contentFirstRow = Math.max(firstRow, titleRows.getLastRow() + 1);
        }
        int contentFirstColumn = firstColumn;
        if (titleColumns != null && titleColumns.getFirstColumn() <= firstColumn) {
            contentFirstColumn = Math.max(firstColumn, titleColumns.getLastColumn() + 1);
        }

        double[] columnWidths = new double[lastColumn + 2];
        double[] columnX = new double[lastColumn + 3];
        for (int c = 0; c <= lastColumn; c++) {
            double width = 0;
            if (!sheet.isColumnHidden(c)) {
                width = Units.pixelsToPoints(sheet.getColumnWidthInPixels(c));
            }
            if (width <= 0.01 && !sheet.isColumnHidden(c)) {
                width = Units.pixelsToPoints(sheet.getDefaultColumnWidth() * 7.0);
            }
            columnWidths[c] = width * columnWidthScale;
            columnX[c + 1] = columnX[c] + columnWidths[c];
        }
        // Ensure the array has a value for one column past the print area.
        columnX[lastColumn + 2] = columnX[lastColumn + 1];

        double defaultRowHeight = sheet.getDefaultRowHeightInPoints();
        double[] rowHeights = new double[lastRow + 2];
        double[] rowY = new double[lastRow + 3];
        for (int r = 0; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            double height = row == null ? defaultRowHeight : row.getHeightInPoints();
            if (row != null && row.getZeroHeight()) {
                height = 0;
            }
            rowHeights[r] = Math.max(0, height);
            rowY[r + 1] = rowY[r] + rowHeights[r];
        }
        rowY[lastRow + 2] = rowY[lastRow + 1];
        double[] renderRowY = new double[rowY.length];
        for (int r = 0; r < rowY.length; r++) {
            renderRowY[r] = rowY[r] * ROW_HEIGHT_SCALE;
        }

        PrintSettingsModel print = readPrintSettings(workbook, sheet);

        return new SheetLayout(workbook, sheet, sheetIndex,
                firstRow, lastRow, firstColumn, lastColumn,
                contentFirstRow, contentFirstColumn,
                titleFirstRow, titleLastRow, titleFirstColumn, titleLastColumn,
                columnWidths, columnX, rowHeights, rowY, renderRowY, print,
                new ArrayList<>(sheet.getMergedRegions()));
    }

    private static int[] readPrintArea(XSSFWorkbook workbook, int sheetIndex, XSSFSheet sheet) {
        String reference = workbook.getPrintArea(sheetIndex);
        if (reference != null && !reference.isBlank()) {
            String area = reference.substring(reference.lastIndexOf('!') + 1);
            try {
                AreaReference ref = new AreaReference(area, SpreadsheetVersion.EXCEL2007);
                return new int[]{
                        ref.getFirstCell().getRow(),
                        ref.getLastCell().getRow(),
                        ref.getFirstCell().getCol(),
                        ref.getLastCell().getCol()
                };
            } catch (RuntimeException ignored) {
                // Fall through to the used range.
            }
        }
        int firstRow = Math.max(0, sheet.getFirstRowNum());
        int lastRow = Math.max(firstRow, sheet.getLastRowNum());
        int firstColumn = Integer.MAX_VALUE;
        int lastColumn = 0;
        for (int r = firstRow; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            if (row.getFirstCellNum() >= 0) {
                firstColumn = Math.min(firstColumn, row.getFirstCellNum());
            }
            if (row.getLastCellNum() > 0) {
                lastColumn = Math.max(lastColumn, row.getLastCellNum() - 1);
            }
        }
        if (firstColumn == Integer.MAX_VALUE) {
            firstColumn = 0;
        }
        return new int[]{firstRow, lastRow, firstColumn, lastColumn};
    }

    private static PrintSettingsModel readPrintSettings(XSSFWorkbook workbook, XSSFSheet sheet) {
        PrintSetup setup = sheet.getPrintSetup();
        int code = setup.getPaperSize();
        double[] size = Units.paperSizePoints(code);
        boolean landscape = setup.getLandscape();
        double paperWidth = landscape ? size[1] : size[0];
        double paperHeight = landscape ? size[0] : size[1];

        double scale = setup.getScale() > 0 ? setup.getScale() / 100d : 1d;
        int fitWidth = setup.getFitWidth();
        int fitHeight = setup.getFitHeight();
        boolean fitToPage = sheet.getFitToPage();

        int firstPageNumber = setup.getPageStart() & 0xFFFF;
        if (firstPageNumber == 0xFFFE || firstPageNumber == 0) {
            firstPageNumber = -1;
        }

        Header header = sheet.getHeader();
        Footer footer = sheet.getFooter();
        HeaderFooterModel headerModel = new HeaderFooterModel(
                safe(header == null ? null : header.getLeft()),
                safe(header == null ? null : header.getCenter()),
                safe(header == null ? null : header.getRight()));
        HeaderFooterModel footerModel = new HeaderFooterModel(
                safe(footer == null ? null : footer.getLeft()),
                safe(footer == null ? null : footer.getCenter()),
                safe(footer == null ? null : footer.getRight()));

        List<Integer> rowBreaks = new ArrayList<>();
        for (int b : sheet.getRowBreaks()) {
            rowBreaks.add(b);
        }
        List<Integer> columnBreaks = new ArrayList<>();
        for (int b : sheet.getColumnBreaks()) {
            columnBreaks.add(b);
        }
        rowBreaks.sort(Integer::compareTo);
        columnBreaks.sort(Integer::compareTo);

        return new PrintSettingsModel(
                code,
                paperWidth,
                paperHeight,
                landscape,
                Units.inchesToPoints(sheet.getMargin(PageMargin.LEFT)),
                Units.inchesToPoints(sheet.getMargin(PageMargin.RIGHT)),
                Units.inchesToPoints(sheet.getMargin(PageMargin.TOP)),
                Units.inchesToPoints(sheet.getMargin(PageMargin.BOTTOM)),
                Units.inchesToPoints(sheet.getMargin(PageMargin.HEADER)),
                Units.inchesToPoints(sheet.getMargin(PageMargin.FOOTER)),
                scale,
                fitToPage,
                fitWidth,
                fitHeight,
                sheet.getHorizontallyCenter(),
                sheet.getVerticallyCenter(),
                sheet.isPrintGridlines(),
                setup.getLeftToRight(),
                firstPageNumber,
                setup.getUsePage(),
                sheet.getRepeatingRows(),
                sheet.getRepeatingColumns(),
                rowBreaks,
                columnBreaks,
                headerModel,
                footerModel);
    }

    private static String safe(String value) {
        return value == null ? "" : value;
    }

    public double titleWidth() {
        if (titleFirstColumn < 0) {
            return 0;
        }
        int start = Math.max(0, titleFirstColumn);
        int end = Math.min(lastColumn, titleLastColumn);
        if (end < start) {
            return 0;
        }
        return columnX[end + 1] - columnX[start];
    }

    public double titleHeight() {
        if (titleFirstRow < 0) {
            return 0;
        }
        int start = Math.max(0, titleFirstRow);
        int end = Math.min(lastRow, titleLastRow);
        if (end < start) {
            return 0;
        }
        return rowY[end + 1] - rowY[start];
    }

    public static double renderRowOffset(double offset) {
        return offset * ROW_HEIGHT_SCALE;
    }

    public double renderTitleHeight() {
        if (titleFirstRow < 0) {
            return 0;
        }
        int start = Math.max(0, titleFirstRow);
        int end = Math.min(lastRow, titleLastRow);
        if (end < start) {
            return 0;
        }
        return renderRowY[end + 1] - renderRowY[start];
    }

    public double contentWidth() {
        return Math.max(0, columnX[lastColumn + 1] - columnX[contentFirstColumn]);
    }

    public double contentHeight() {
        return Math.max(0, rowY[lastRow + 1] - rowY[contentFirstRow]);
    }

    public boolean contains(int row, int column) {
        return row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn;
    }

    public int[] sortedRowBreaks() {
        return print.rowBreaks.stream().mapToInt(Integer::intValue).toArray();
    }

    public int[] sortedColumnBreaks() {
        return print.columnBreaks.stream().mapToInt(Integer::intValue).toArray();
    }

    @Override
    public String toString() {
        return "SheetLayout{" +
                "name='" + name + '\'' +
                ", area=" + firstRow + ":" + lastRow + "/" + firstColumn + ":" + lastColumn +
                ", content=" + contentFirstRow + ":" + contentFirstColumn +
                ", titles=" + titleFirstRow + ":" + titleLastRow + "/" + titleFirstColumn + ":" + titleLastColumn +
                ", rows=" + Arrays.toString(rowHeights) +
                '}';
    }
}
