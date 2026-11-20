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

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.within;

/**
 * Regression tests for rows that have wrapped cells but no explicit height.
 * Excel auto-fits these rows; POI reports the worksheet default instead.
 */
class RowAutoFitterTest {

    private static final String LONG_TEXT =
            "営業課長が見積を承認する。100 万円以上は営業部長も承認する。";

    @Test
    void expandsWrappedRowWithoutExplicitHeight() throws Exception {
        Path path = Path.of("testdata", "validation", "screen-list-ja.xlsx");
        try (InputStream in = Files.newInputStream(path);
             XSSFWorkbook workbook = new XSSFWorkbook(in)) {
            XSSFSheet sheet = workbook.getSheetAt(0);
            SheetLayout layout = SheetLayout.read(workbook, 0, sheet, 1.0);
            double defaultHeight = sheet.getDefaultRowHeightInPoints();

            assertThat(layout.rowHeights[5]).isEqualTo(defaultHeight);

            fit(layout);

            assertThat(layout.rowHeights[0])
                    .as("the 16pt merged title needs more than the default height")
                    .isGreaterThan(defaultHeight);
            assertThat(layout.rowHeights[5])
                    .as("C6 wraps into two lines")
                    .isBetween(defaultHeight * 1.5, defaultHeight * 2.0);
            assertThat(layout.rowY[6] - layout.rowY[5])
                    .isCloseTo(layout.rowHeights[5], within(0.000001));
            assertThat(layout.renderRowY[6] - layout.renderRowY[5])
                    .isCloseTo(layout.rowHeights[5] * 17d / 18d, within(0.000001));
            assertThat(layout.rowHeights[3])
                    .as("a one-line wrapped row keeps the default height")
                    .isEqualTo(defaultHeight);
        }
    }

    @Test
    void leavesExplicitRowHeightUnchanged() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Sheet1");
            sheet.setColumnWidth(0, 46 * 256);
            XSSFRow row = sheet.createRow(0);
            row.setHeightInPoints(20);
            XSSFCell cell = row.createCell(0);
            cell.setCellValue(LONG_TEXT);
            XSSFCellStyle style = workbook.createCellStyle();
            style.setWrapText(true);
            cell.setCellStyle(style);

            SheetLayout layout = SheetLayout.read(workbook, 0, sheet, 1.0);
            assertThat(layout.rowHeights[0]).isEqualTo(20.0f);

            fit(layout);

            assertThat(layout.rowHeights[0]).isEqualTo(20.0f);
        }
    }

    @Test
    void keepsHiddenAndShrinkToFitRowsAtTheirCurrentHeight() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Sheet1");
            sheet.setColumnWidth(0, 46 * 256);

            XSSFRow hidden = sheet.createRow(0);
            hidden.setZeroHeight(true);
            XSSFCell hiddenCell = hidden.createCell(0);
            hiddenCell.setCellValue(LONG_TEXT);
            XSSFCellStyle wrapped = workbook.createCellStyle();
            wrapped.setWrapText(true);
            hiddenCell.setCellStyle(wrapped);

            XSSFRow shrink = sheet.createRow(1);
            XSSFCell shrinkCell = shrink.createCell(0);
            shrinkCell.setCellValue(LONG_TEXT);
            XSSFCellStyle shrinkStyle = workbook.createCellStyle();
            shrinkStyle.setWrapText(true);
            shrinkStyle.setShrinkToFit(true);
            shrinkCell.setCellStyle(shrinkStyle);

            SheetLayout layout = SheetLayout.read(workbook, 0, sheet, 1.0);
            double defaultHeight = sheet.getDefaultRowHeightInPoints();
            fit(layout);

            assertThat(layout.rowHeights[0]).isZero();
            assertThat(layout.rowHeights[1]).isEqualTo(defaultHeight);
        }
    }

    private static void fit(SheetLayout layout) {
        new RowAutoFitter(new FontResolver(), new TextLayoutEngine(), new DataFormatter())
                .fit(layout);
    }
}
