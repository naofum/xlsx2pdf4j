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

import java.util.ArrayList;
import java.util.List;

/**
 * Splits a sheet into printable pages.
 */
public final class LayoutEngine {

    private static final double MIN_SCALE = 0.1;
    private static final double MAX_SCALE = 4.0;

    public List<PageLayout> layout(SheetLayout sheet, int startPageNumber) {
        return layout(sheet, startPageNumber, 0);
    }

    /**
     * @param printScaleOverride explicit scale, or 0 to use the workbook setting
     */
    public List<PageLayout> layout(SheetLayout sheet, int startPageNumber,
                                   double printScaleOverride) {
        PrintSettingsModel print = sheet.print;
        double titleWidth = sheet.titleWidth();
        double titleHeight = sheet.titleHeight();
        double contentWidth = sheet.contentWidth();
        double contentHeight = sheet.contentHeight();
        if (contentWidth <= 0 || contentHeight <= 0) {
            // Empty sheet: one blank page so the workbook still produces output.
            double emptyScale = printScaleOverride > 0
                    ? Units.clamp(printScaleOverride, MIN_SCALE, MAX_SCALE) : 1d;
            return List.of(new PageLayout(sheet, startPageNumber, sheet.contentFirstRow, sheet.contentFirstRow,
                    sheet.contentFirstColumn, sheet.contentFirstColumn, emptyScale,
                    print.paperWidth, print.paperHeight,
                    print.marginLeft, print.marginTop, 0, 0));
        }

        double totalWidth = contentWidth + (titleWidth > 0 ? titleWidth : 0);
        double totalHeight = contentHeight + (titleHeight > 0 ? titleHeight : 0);
        double usableWidth = print.usableWidth();
        double usableHeight = print.usableHeight();

        double scale;
        if (printScaleOverride > 0) {
            scale = printScaleOverride;
        } else if (print.fitToPage) {
            int fitWidth = print.fitWidth > 0 ? print.fitWidth : 1;
            int fitHeight = print.fitHeight;
            double scaleW = fitWidth > 0 && totalWidth > 0 ? usableWidth / totalWidth : Double.MAX_VALUE;
            double scaleH = fitHeight > 0 && totalHeight > 0 ? usableHeight / totalHeight : Double.MAX_VALUE;
            scale = Math.min(scaleW, scaleH);
            if (!Double.isFinite(scale) || scale <= 0) {
                scale = 1d;
            }
            // Excel's fit-to-page does not normally enlarge content in these
            // templates; capping at 100% also keeps pagination stable.
            scale = Math.min(1.0, scale);
        } else {
            scale = print.scale;
        }
        scale = Units.clamp(scale, MIN_SCALE, MAX_SCALE);

        // Excel allows columns to use the full paper width before inserting a
        // vertical page break; horizontal margins mainly position the content.
        double pageWidthInSheet = print.paperWidth / scale;
        double pageHeightInSheet = usableHeight / scale;
        double contentCapacityWidth = Math.max(1, pageWidthInSheet - titleWidth);
        double contentCapacityHeight = Math.max(1, pageHeightInSheet - titleHeight);

        List<Range> columnPages = splitColumns(sheet, contentCapacityWidth);
        // Excel allows a small amount of overflow before starting a new page.
        List<Range> rowPages = splitRows(sheet, contentCapacityHeight * 1.05);
        if (columnPages.isEmpty()) {
            columnPages.add(new Range(sheet.contentFirstColumn, sheet.contentFirstColumn));
        }
        if (rowPages.isEmpty()) {
            rowPages.add(new Range(sheet.contentFirstRow, sheet.contentFirstRow));
        }

        double fullWidth = (sheet.columnX[sheet.lastColumn + 1] - sheet.columnX[sheet.contentFirstColumn]) * scale;
        double fullHeight = (sheet.renderRowY[sheet.lastRow + 1]
                - sheet.renderRowY[sheet.contentFirstRow]) * scale;
        double centerX = print.centerHorizontally
                ? Math.max(0, (usableWidth - (titleWidth * scale + fullWidth)) / 2d)
                : 0;
        double centerY = print.centerVertically
                ? Math.max(0, (usableHeight - (titleHeight * scale + fullHeight)) / 2d)
                : 0;

        List<PageLayout> pages = new ArrayList<>();
        int pageIndex = 0;
        if (print.overThenDown) {
            for (Range rows : rowPages) {
                for (Range columns : columnPages) {
                    pages.add(page(sheet, startPageNumber + pageIndex++, rows, columns, scale,
                            print, titleWidth, titleHeight, centerX, centerY));
                }
            }
        } else {
            for (Range columns : columnPages) {
                for (Range rows : rowPages) {
                    pages.add(page(sheet, startPageNumber + pageIndex++, rows, columns, scale,
                            print, titleWidth, titleHeight, centerX, centerY));
                }
            }
        }
        return pages;
    }

    private PageLayout page(SheetLayout sheet, int pageNumber, Range rows, Range columns, double scale,
                            PrintSettingsModel print, double titleWidth, double titleHeight,
                            double centerX, double centerY) {
        double contentLeft = print.marginLeft + titleWidth * scale + centerX;
        double contentTop = print.marginTop + sheet.renderTitleHeight() * scale + centerY;
        double width = (sheet.columnX[columns.last + 1] - sheet.columnX[columns.first]) * scale;
        double height = (sheet.renderRowY[rows.last + 1] - sheet.renderRowY[rows.first]) * scale;
        return new PageLayout(sheet, pageNumber, rows.first, rows.last, columns.first, columns.last, scale,
                print.paperWidth, print.paperHeight, contentLeft, contentTop, width, height);
    }

    private List<Range> splitColumns(SheetLayout sheet, double capacity) {
        List<Range> pages = new ArrayList<>();
        int start = sheet.contentFirstColumn;
        int current = start;
        double used = 0;
        while (current <= sheet.lastColumn) {
            if (current > start && containsBreak(sheet.print.columnBreaks, current - 1)) {
                pages.add(new Range(start, current - 1));
                start = current;
                used = 0;
            }
            double width = sheet.columnWidths[current];
            if (current > start && used + width > capacity) {
                pages.add(new Range(start, current - 1));
                start = current;
                used = 0;
            }
            used += width;
            current++;
        }
        if (start <= sheet.lastColumn) {
            pages.add(new Range(start, sheet.lastColumn));
        }
        return pages;
    }

    private List<Range> splitRows(SheetLayout sheet, double capacity) {
        List<Range> pages = new ArrayList<>();
        int start = sheet.contentFirstRow;
        int current = start;
        double used = 0;
        while (current <= sheet.lastRow) {
            if (current > start && containsBreak(sheet.print.rowBreaks, current - 1)) {
                pages.add(new Range(start, current - 1));
                start = current;
                used = 0;
            }
            double height = sheet.rowHeights[current];
            if (current > start && used + height > capacity) {
                pages.add(new Range(start, current - 1));
                start = current;
                used = 0;
            }
            used += height;
            current++;
        }
        if (start <= sheet.lastRow) {
            pages.add(new Range(start, sheet.lastRow));
        }
        return pages;
    }

    private boolean containsBreak(List<Integer> breaks, int rowOrColumn) {
        for (int value : breaks) {
            if (value == rowOrColumn) {
                return true;
            }
        }
        return false;
    }

    private record Range(int first, int last) {
    }
}
