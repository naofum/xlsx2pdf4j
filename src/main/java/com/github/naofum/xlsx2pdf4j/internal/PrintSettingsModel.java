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

import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * Normalized print settings for one sheet. Lengths are in points.
 */
public final class PrintSettingsModel {

    public final int paperSizeCode;
    public final double paperWidth;
    public final double paperHeight;
    public final boolean landscape;
    public final double marginLeft;
    public final double marginRight;
    public final double marginTop;
    public final double marginBottom;
    public final double headerMargin;
    public final double footerMargin;
    public final double scale;
    public final boolean fitToPage;
    public final int fitWidth;
    public final int fitHeight;
    public final boolean centerHorizontally;
    public final boolean centerVertically;
    public final boolean printGridlines;
    public final boolean overThenDown;
    public final int firstPageNumber;
    public final boolean useFirstPageNumber;
    public final CellRangeAddress titleRows;
    public final CellRangeAddress titleColumns;
    public final List<Integer> rowBreaks;
    public final List<Integer> columnBreaks;
    public final HeaderFooterModel header;
    public final HeaderFooterModel footer;

    public PrintSettingsModel(int paperSizeCode,
                              double paperWidth,
                              double paperHeight,
                              boolean landscape,
                              double marginLeft,
                              double marginRight,
                              double marginTop,
                              double marginBottom,
                              double headerMargin,
                              double footerMargin,
                              double scale,
                              boolean fitToPage,
                              int fitWidth,
                              int fitHeight,
                              boolean centerHorizontally,
                              boolean centerVertically,
                              boolean printGridlines,
                              boolean overThenDown,
                              int firstPageNumber,
                              boolean useFirstPageNumber,
                              CellRangeAddress titleRows,
                              CellRangeAddress titleColumns,
                              List<Integer> rowBreaks,
                              List<Integer> columnBreaks,
                              HeaderFooterModel header,
                              HeaderFooterModel footer) {
        this.paperSizeCode = paperSizeCode;
        this.paperWidth = paperWidth;
        this.paperHeight = paperHeight;
        this.landscape = landscape;
        this.marginLeft = marginLeft;
        this.marginRight = marginRight;
        this.marginTop = marginTop;
        this.marginBottom = marginBottom;
        this.headerMargin = headerMargin;
        this.footerMargin = footerMargin;
        this.scale = scale;
        this.fitToPage = fitToPage;
        this.fitWidth = fitWidth;
        this.fitHeight = fitHeight;
        this.centerHorizontally = centerHorizontally;
        this.centerVertically = centerVertically;
        this.printGridlines = printGridlines;
        this.overThenDown = overThenDown;
        this.firstPageNumber = firstPageNumber;
        this.useFirstPageNumber = useFirstPageNumber;
        this.titleRows = titleRows;
        this.titleColumns = titleColumns;
        this.rowBreaks = List.copyOf(rowBreaks);
        this.columnBreaks = List.copyOf(columnBreaks);
        this.header = header;
        this.footer = footer;
    }

    public double usableWidth() {
        return Math.max(1, paperWidth - marginLeft - marginRight);
    }

    public double usableHeight() {
        return Math.max(1, paperHeight - marginTop - marginBottom);
    }
}
