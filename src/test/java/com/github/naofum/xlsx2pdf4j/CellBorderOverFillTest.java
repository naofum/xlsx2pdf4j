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

import com.github.naofum.xlsx2pdf4j.api.ConversionOptions;
import com.github.naofum.xlsx2pdf4j.api.Xlsx2Pdf;
import com.github.naofum.xlsx2pdf4j.api.TextLayerMode;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.awt.image.BufferedImage;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

class CellBorderOverFillTest {

    @Test
    void bottomBorderIsNotCoveredByTheNextRowsFill() throws Exception {
        Path workbookPath = Path.of("target", "cell-border-over-fill.xlsx");
        Path pdfPath = Path.of("target", "cell-border-over-fill.pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Border");
            sheet.setColumnWidth(0, 30 * 256);

            Row top = sheet.createRow(0);
            top.setHeightInPoints(24);
            Cell topCell = top.createCell(0);
            CellStyle border = workbook.createCellStyle();
            border.setBorderBottom(BorderStyle.THIN);
            border.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            topCell.setCellStyle(border);

            Row bottom = sheet.createRow(1);
            bottom.setHeightInPoints(24);
            Cell bottomCell = bottom.createCell(0);
            CellStyle fill = workbook.createCellStyle();
            fill.setFillForegroundColor(IndexedColors.WHITE.getIndex());
            fill.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            bottomCell.setCellStyle(fill);

            workbook.setPrintArea(0, "A1:A2");
            try (OutputStream out = Files.newOutputStream(workbookPath)) {
                workbook.write(out);
            }
        }

        Xlsx2Pdf.convert(workbookPath, pdfPath,
                ConversionOptions.builder()
                        .dpi(150)
                        .textLayerMode(TextLayerMode.NONE)
                        .build());

        try (PDDocument document = Loader.loadPDF(pdfPath.toFile())) {
            assertThat(longestDarkRun(document, 0))
                    .as("the bottom border must remain visible over the next row's white fill")
                    .isGreaterThan(100);
        }
    }

    @Test
    void borderReferencedWithoutApplyBorderIsRendered() throws Exception {
        Path input = Path.of("testdata", "validation", "screen-list-ja.xlsx");
        Path pdfPath = Path.of("target", "screen-list-borders.pdf");

        Xlsx2Pdf.convert(input, pdfPath,
                ConversionOptions.builder()
                        .dpi(150)
                        .textLayerMode(TextLayerMode.NONE)
                        .build());

        try (PDDocument document = Loader.loadPDF(pdfPath.toFile())) {
            assertThat(longestDarkRun(document, 0))
                    .as("a borderId without applyBorder must still be drawn")
                    .isGreaterThan(300);
        }
    }

    @Test
    void previousRowsBottomBorderIsDrawnAtTheTopOfTheNextPage() throws Exception {
        Path workbookPath = Path.of("target", "page-break-border.xlsx");
        Path pdfPath = Path.of("target", "page-break-border.pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Border");
            sheet.setColumnWidth(0, 30 * 256);

            CellStyle border = workbook.createCellStyle();
            border.setBorderBottom(BorderStyle.THIN);
            border.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            for (int r = 0; r <= 6; r++) {
                Row row = sheet.createRow(r);
                row.setHeightInPoints(24);
                row.createCell(0).setCellStyle(border);
            }

            Row firstOnNextPage = sheet.createRow(7);
            firstOnNextPage.setHeightInPoints(24);
            firstOnNextPage.createCell(0).setCellStyle(workbook.createCellStyle());

            sheet.setRowBreak(6);
            workbook.setPrintArea(0, "A1:A8");
            try (OutputStream out = Files.newOutputStream(workbookPath)) {
                workbook.write(out);
            }
        }

        Xlsx2Pdf.convert(workbookPath, pdfPath,
                ConversionOptions.builder()
                        .dpi(150)
                        .textLayerMode(TextLayerMode.NONE)
                        .build());

        try (PDDocument document = Loader.loadPDF(pdfPath.toFile())) {
            assertThat(document.getNumberOfPages()).isGreaterThanOrEqualTo(2);
            assertThat(longestDarkRun(document, 1))
                    .as("the previous row's bottom border must continue at the page top")
                    .isGreaterThan(100);
        }
    }

    private static int longestDarkRun(PDDocument document, int page) throws Exception {
        BufferedImage image = new PDFRenderer(document)
                .renderImageWithDPI(page, 150, ImageType.RGB);
        int longest = 0;
        for (int y = 0; y < image.getHeight(); y++) {
            int run = 0;
            for (int x = 0; x < image.getWidth(); x++) {
                int rgb = image.getRGB(x, y);
                int luminance = (((rgb >> 16) & 0xFF) + ((rgb >> 8) & 0xFF) + (rgb & 0xFF)) / 3;
                if (luminance < 128) {
                    run++;
                    longest = Math.max(longest, run);
                } else {
                    run = 0;
                }
            }
        }
        return longest;
    }
}
