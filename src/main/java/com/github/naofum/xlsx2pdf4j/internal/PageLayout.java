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

/**
 * One printed page. Rows and columns are absolute sheet coordinates.
 */
public final class PageLayout {

    public final SheetLayout sheet;
    public final int pageNumber;
    public final int firstRow;
    public final int lastRow;
    public final int firstColumn;
    public final int lastColumn;
    public final double scale;
    public final double pageWidth;
    public final double pageHeight;
    public final double contentLeft;
    public final double contentTop;
    public final double contentWidth;
    public final double contentHeight;

    public PageLayout(SheetLayout sheet,
                      int pageNumber,
                      int firstRow,
                      int lastRow,
                      int firstColumn,
                      int lastColumn,
                      double scale,
                      double pageWidth,
                      double pageHeight,
                      double contentLeft,
                      double contentTop,
                      double contentWidth,
                      double contentHeight) {
        this.sheet = sheet;
        this.pageNumber = pageNumber;
        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstColumn = firstColumn;
        this.lastColumn = lastColumn;
        this.scale = scale;
        this.pageWidth = pageWidth;
        this.pageHeight = pageHeight;
        this.contentLeft = contentLeft;
        this.contentTop = contentTop;
        this.contentWidth = contentWidth;
        this.contentHeight = contentHeight;
    }
}
