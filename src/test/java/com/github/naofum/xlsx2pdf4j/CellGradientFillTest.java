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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.awt.Rectangle;
import java.awt.image.BufferedImage;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Cell gradient fills live in styles.xml, but poi-ooxml-lite does not expose
 * the gradient schema type. The tests write a normal solid fill, then replace
 * its XML with a gradient before conversion.
 */
class CellGradientFillTest {

    private static final String SOLID_FILL =
            "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"123456\"/></patternFill></fill>";

    @Test
    void linearGradientRunsLeftToRight() throws Exception {
        BufferedImage page = convert("linear", """
                <fill><gradientFill degree="0">
                  <stop position="0"><color rgb="FFFF0000"/></stop>
                  <stop position="1"><color rgb="FF0000FF"/></stop>
                </gradientFill></fill>
                """);
        Rectangle bounds = gradientBounds(page);

        int left = sample(page, bounds, 0.08, 0.5);
        int right = sample(page, bounds, 0.92, 0.5);
        assertThat(red(left)).as("left stop is red").isGreaterThan(blue(left) + 80);
        assertThat(blue(right)).as("right stop is blue").isGreaterThan(red(right) + 80);
    }

    @Test
    void linearGradientFollowsDegree() throws Exception {
        BufferedImage page = convert("vertical", """
                <fill><gradientFill degree="90">
                  <stop position="0"><color rgb="FFFF0000"/></stop>
                  <stop position="1"><color rgb="FF0000FF"/></stop>
                </gradientFill></fill>
                """);
        Rectangle bounds = gradientBounds(page);

        int top = sample(page, bounds, 0.5, 0.08);
        int bottom = sample(page, bounds, 0.5, 0.92);
        assertThat(red(top)).as("top stop is red").isGreaterThan(blue(top) + 80);
        assertThat(blue(bottom)).as("bottom stop is blue").isGreaterThan(red(bottom) + 80);
    }

    @Test
    void multiStopGradientInterpolatesMiddleStop() throws Exception {
        BufferedImage page = convert("multi", """
                <fill><gradientFill degree="0">
                  <stop position="0"><color rgb="FFFF0000"/></stop>
                  <stop position="0.5"><color rgb="FFFFFF00"/></stop>
                  <stop position="1"><color rgb="FF00B050"/></stop>
                </gradientFill></fill>
                """);
        Rectangle bounds = gradientBounds(page);

        int middle = sample(page, bounds, 0.5, 0.5);
        assertThat(red(middle)).as("middle stop has red").isGreaterThan(180);
        assertThat(green(middle)).as("middle stop has green").isGreaterThan(180);
        assertThat(blue(middle)).as("middle stop has little blue").isLessThan(80);
        int end = sample(page, bounds, 0.95, 0.5);
        assertThat(green(end)).as("end stop is green").isGreaterThan(red(end) + 80);
    }

    @Test
    void pathGradientRadiatesFromCellCentre() throws Exception {
        BufferedImage page = convert("path", """
                <fill><gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5">
                  <stop position="0"><color rgb="FFFFFFFF"/></stop>
                  <stop position="1"><color rgb="FF7030A0"/></stop>
                </gradientFill></fill>
                """);
        Rectangle bounds = gradientBounds(page);

        int centre = sample(page, bounds, 0.5, 0.5);
        int edge = sample(page, bounds, 0.03, 0.5);
        assertThat(luminance(centre))
                .as("path gradient centre is white")
                .isGreaterThan(220);
        assertThat(luminance(edge))
                .as("path gradient edge is purple")
                .isLessThan(luminance(centre) - 50);
    }

    @Test
    void gradientSpansMergedCell() throws Exception {
        BufferedImage page = convert("merged", """
                <fill><gradientFill degree="0">
                  <stop position="0"><color rgb="FFFF0000"/></stop>
                  <stop position="1"><color rgb="FF0000FF"/></stop>
                </gradientFill></fill>
                """, true);
        Rectangle bounds = gradientBounds(page);

        assertThat(bounds.width)
                .as("gradient covers all merged columns")
                .isGreaterThan(250);
        int left = sample(page, bounds, 0.05, 0.5);
        int right = sample(page, bounds, 0.95, 0.5);
        assertThat(red(left)).isGreaterThan(blue(left) + 80);
        assertThat(blue(right)).isGreaterThan(red(right) + 80);
    }

    private BufferedImage convert(String name, String gradientFill) throws Exception {
        return convert(name, gradientFill, false);
    }

    private BufferedImage convert(String name, String gradientFill, boolean merged) throws Exception {
        Path base = Path.of("target", "cell-gradient-" + name + "-base.xlsx");
        Path gradient = Path.of("target", "cell-gradient-" + name + ".xlsx");
        Path pdf = Path.of("target", "cell-gradient-" + name + ".pdf");
        Files.createDirectories(base.getParent());
        createBaseWorkbook(base, merged);
        replaceFill(base, gradient, gradientFill);
        Xlsx2Pdf.convert(gradient, pdf,
                ConversionOptions.builder()
                        .dpi(150)
                        .textLayerMode(TextLayerMode.NONE)
                        .build());
        try (PDDocument document = Loader.loadPDF(pdf.toFile())) {
            assertThat(document.getNumberOfPages()).isEqualTo(1);
            return new PDFRenderer(document).renderImageWithDPI(0, 150, ImageType.RGB);
        }
    }

    private void createBaseWorkbook(Path path, boolean merged) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Gradient");
            sheet.setColumnWidth(0, 24 * 256);
            if (merged) {
                sheet.setColumnWidth(1, 24 * 256);
                sheet.setColumnWidth(2, 24 * 256);
                sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));
            }
            Row row = sheet.createRow(0);
            row.setHeightInPoints(36);
            Cell cell = row.createCell(0);
            cell.setCellValue(" ");

            XSSFCellStyle style = workbook.createCellStyle();
            style.setFillForegroundColor(new XSSFColor(new byte[]{0x12, 0x34, 0x56}, null));
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(style);
            workbook.setPrintArea(0, merged ? "A1:C1" : "A1:A1");
            try (OutputStream out = Files.newOutputStream(path)) {
                workbook.write(out);
            }
        }
    }

    private void replaceFill(Path source, Path target, String gradientFill) throws Exception {
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
                        assertThat(xml).contains(SOLID_FILL);
                        xml = xml.replace(SOLID_FILL, gradientFill.strip());
                        bytes = xml.getBytes(StandardCharsets.UTF_8);
                    }
                    out.write(bytes);
                }
                out.closeEntry();
            }
        }
    }

    private static Rectangle gradientBounds(BufferedImage image) {
        int minX = Integer.MAX_VALUE;
        int minY = Integer.MAX_VALUE;
        int maxX = Integer.MIN_VALUE;
        int maxY = Integer.MIN_VALUE;
        for (int y = 0; y < image.getHeight(); y++) {
            for (int x = 0; x < image.getWidth(); x++) {
                int rgb = image.getRGB(x, y);
                if (red(rgb) < 245 || green(rgb) < 245 || blue(rgb) < 245) {
                    minX = Math.min(minX, x);
                    minY = Math.min(minY, y);
                    maxX = Math.max(maxX, x);
                    maxY = Math.max(maxY, y);
                }
            }
        }
        assertThat(maxX).isGreaterThan(minX);
        assertThat(maxY).isGreaterThan(minY);
        return new Rectangle(minX, minY, maxX - minX + 1, maxY - minY + 1);
    }

    private static int sample(BufferedImage image, Rectangle bounds, double fx, double fy) {
        int x = bounds.x + (int) Math.round(fx * (bounds.width - 1));
        int y = bounds.y + (int) Math.round(fy * (bounds.height - 1));
        return image.getRGB(x, y);
    }

    private static int red(int rgb) {
        return (rgb >> 16) & 0xFF;
    }

    private static int green(int rgb) {
        return (rgb >> 8) & 0xFF;
    }

    private static int blue(int rgb) {
        return rgb & 0xFF;
    }

    private static int luminance(int rgb) {
        return (red(rgb) + green(rgb) + blue(rgb)) / 3;
    }
}
