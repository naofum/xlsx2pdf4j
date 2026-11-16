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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.awt.image.BufferedImage;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

class ZeroHeightLineShapeTest {

    @Test
    void drawsHorizontalLineShapeWithZeroHeight() throws Exception {
        Path workbookPath = Path.of("target", "zero-height-line.xlsx");
        Path pdfPath = Path.of("target", "zero-height-line.pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Line");
            for (int c = 0; c < 4; c++) {
                sheet.setColumnWidth(c, 20 * 256);
            }
            Row row = sheet.createRow(0);
            row.setHeightInPoints(24);

            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 0, 3, 0);
            XSSFSimpleShape line = drawing.createSimpleShape(anchor);
            line.setShapeType(ShapeTypes.LINE);

            workbook.setPrintArea(0, "A1:D1");
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
            BufferedImage image = new PDFRenderer(document)
                    .renderImageWithDPI(0, 150, ImageType.RGB);
            int longest = 0;
            for (int y = 0; y < image.getHeight(); y++) {
                int run = 0;
                for (int x = 0; x < image.getWidth(); x++) {
                    int rgb = image.getRGB(x, y);
                    int luminance = (((rgb >> 16) & 0xFF)
                            + ((rgb >> 8) & 0xFF) + (rgb & 0xFF)) / 3;
                    if (luminance < 128) {
                        run++;
                        longest = Math.max(longest, run);
                    } else {
                        run = 0;
                    }
                }
            }
            assertThat(longest)
                    .as("a horizontal line shape with zero height must still be drawn")
                    .isGreaterThan(50);
        }
    }
}
