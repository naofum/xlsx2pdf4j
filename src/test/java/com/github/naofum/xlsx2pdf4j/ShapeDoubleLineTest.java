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
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.main.STCompoundLine;

import java.awt.image.BufferedImage;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class ShapeDoubleLineTest {

    @Test
    void doubleCompoundLineDrawsTwoParallelBorders() throws Exception {
        Path workbookPath = Path.of("target", "shape-double-line.xlsx");
        Path pdfPath = Path.of("target", "shape-double-line.pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Line");
            sheet.createRow(0).createCell(0).setCellValue(" ");
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            XSSFSimpleShape shape = drawing.createSimpleShape(
                    new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 7, 7));
            shape.setShapeType(ShapeTypes.RECT);
            shape.setFillColor(0xFF, 0xFF, 0xFF);
            shape.setLineStyleColor(0, 0, 0);
            shape.setLineWidth(9);
            shape.getCTShape().getSpPr().getLn().setCmpd(STCompoundLine.DBL);

            workbook.setPrintArea(0, "A1:H8");
            try (OutputStream out = Files.newOutputStream(workbookPath)) {
                workbook.write(out);
            }
        }

        Xlsx2Pdf.convert(workbookPath, pdfPath,
                ConversionOptions.builder()
                        .dpi(150)
                        .textLayerMode(TextLayerMode.NONE)
                        .build());

        BufferedImage page;
        try (PDDocument document = Loader.loadPDF(pdfPath.toFile())) {
            assertThat(document.getNumberOfPages()).isEqualTo(1);
            page = new PDFRenderer(document).renderImageWithDPI(0, 150, ImageType.RGB);
        }

        int[] bounds = darkBounds(page);
        int y = (bounds[1] + bounds[3]) / 2;
        List<Integer> starts = new ArrayList<>();
        boolean previousDark = false;
        for (int x = bounds[0]; x <= bounds[2]; x++) {
            boolean dark = isDark(page.getRGB(x, y));
            if (dark && !previousDark) {
                starts.add(x);
            }
            previousDark = dark;
        }
        assertThat(starts)
                .as("a double outline has two dark runs on each side")
                .hasSizeGreaterThanOrEqualTo(4);
    }

    private static int[] darkBounds(BufferedImage image) {
        int minX = Integer.MAX_VALUE;
        int minY = Integer.MAX_VALUE;
        int maxX = Integer.MIN_VALUE;
        int maxY = Integer.MIN_VALUE;
        for (int y = 0; y < image.getHeight(); y++) {
            for (int x = 0; x < image.getWidth(); x++) {
                if (isDark(image.getRGB(x, y))) {
                    minX = Math.min(minX, x);
                    minY = Math.min(minY, y);
                    maxX = Math.max(maxX, x);
                    maxY = Math.max(maxY, y);
                }
            }
        }
        assertThat(maxX).isGreaterThan(minX);
        assertThat(maxY).isGreaterThan(minY);
        return new int[]{minX, minY, maxX, maxY};
    }

    private static boolean isDark(int rgb) {
        return ((rgb >> 16) & 0xFF) < 90
                && ((rgb >> 8) & 0xFF) < 90
                && (rgb & 0xFF) < 90;
    }
}
