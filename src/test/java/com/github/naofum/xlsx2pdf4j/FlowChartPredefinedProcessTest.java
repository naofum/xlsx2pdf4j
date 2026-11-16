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
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.awt.Rectangle;
import java.awt.image.BufferedImage;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * flowChartPredefinedProcess is a rectangle with two inner vertical lines at
 * one eighth and seven eighths of the shape width.
 */
class FlowChartPredefinedProcessTest {

    @Test
    void predefinedProcessDrawsInnerVerticalLines() throws Exception {
        Path workbookPath = Path.of("target", "predefined-process.xlsx");
        Path pdfPath = Path.of("target", "predefined-process.pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Predefined");
            sheet.createRow(0).createCell(0).setCellValue(" ");
            for (int c = 0; c < 7; c++) {
                sheet.setColumnWidth(c, 12 * 256);
            }
            for (int r = 0; r < 7; r++) {
                if (sheet.getRow(r) == null) {
                    sheet.createRow(r);
                }
                sheet.getRow(r).setHeightInPoints(20);
            }
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            XSSFSimpleShape shape = drawing.createSimpleShape(
                    new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 6, 6));
            shape.setShapeType(ShapeTypes.FLOW_CHART_PREDEFINED_PROCESS);
            shape.setFillColor(0xFF, 0x99, 0xCC);
            shape.setLineStyleColor(0, 0, 0);
            shape.setLineWidth(1.0);
            workbook.setPrintArea(0, "A1:G7");
            try (var out = Files.newOutputStream(workbookPath)) {
                workbook.write(out);
            }
        }

        Xlsx2Pdf.convert(workbookPath, pdfPath,
                ConversionOptions.builder()
                        .dpi(150)
                        .textLayerMode(TextLayerMode.NONE)
                        .build());
        try (PDDocument document = Loader.loadPDF(pdfPath.toFile())) {
            BufferedImage image = new PDFRenderer(document)
                    .renderImageWithDPI(0, 150, ImageType.RGB);
            Rectangle bounds = pinkBounds(image);
            assertThat(bounds).as("predefined process should be drawn").isNotNull();

            int leftLine = darkPixelsNear(image, bounds,
                    bounds.x + bounds.width / 8.0);
            int rightLine = darkPixelsNear(image, bounds,
                    bounds.x + bounds.width * 7.0 / 8.0);
            assertThat(leftLine)
                    .as("left inner vertical line")
                    .isGreaterThan(10);
            assertThat(rightLine)
                    .as("right inner vertical line")
                    .isGreaterThan(10);
        }
    }

    private static int darkPixelsNear(BufferedImage image, Rectangle bounds, double expectedX) {
        int count = 0;
        int fromX = (int) Math.round(expectedX) - 2;
        int toX = (int) Math.round(expectedX) + 2;
        int fromY = bounds.y + (int) (bounds.height * 0.2);
        int toY = bounds.y + (int) (bounds.height * 0.8);
        for (int y = fromY; y <= toY; y++) {
            for (int x = fromX; x <= toX; x++) {
                if (x < 0 || y < 0 || x >= image.getWidth() || y >= image.getHeight()) {
                    continue;
                }
                int rgb = image.getRGB(x, y);
                int red = (rgb >> 16) & 0xFF;
                int green = (rgb >> 8) & 0xFF;
                int blue = rgb & 0xFF;
                if (red < 100 && green < 100 && blue < 100) {
                    count++;
                }
            }
        }
        return count;
    }

    private static Rectangle pinkBounds(BufferedImage image) {
        int minX = Integer.MAX_VALUE;
        int minY = Integer.MAX_VALUE;
        int maxX = -1;
        int maxY = -1;
        for (int y = 0; y < image.getHeight(); y++) {
            for (int x = 0; x < image.getWidth(); x++) {
                int rgb = image.getRGB(x, y);
                int red = (rgb >> 16) & 0xFF;
                int green = (rgb >> 8) & 0xFF;
                int blue = rgb & 0xFF;
                if (red > 235 && green > 120 && green < 205 && blue > 175) {
                    minX = Math.min(minX, x);
                    minY = Math.min(minY, y);
                    maxX = Math.max(maxX, x);
                    maxY = Math.max(maxY, y);
                }
            }
        }
        return maxX < 0 ? null : new Rectangle(minX, minY,
                maxX - minX + 1, maxY - minY + 1);
    }
}
