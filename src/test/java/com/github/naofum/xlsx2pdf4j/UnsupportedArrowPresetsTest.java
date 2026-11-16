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
import com.github.naofum.xlsx2pdf4j.api.ConversionResult;
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

import java.awt.image.BufferedImage;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Covers the arrow presets that were previously approximated by a rectangle.
 * Each page contains one filled preset; a rectangle approximation would cover
 * almost the entire tile, while the real arrow leaves white corners.
 */
class UnsupportedArrowPresetsTest {

    @Test
    void allUnsupportedArrowPresetsRenderAsArrows() throws Exception {
        Map<String, Integer> shapes = new LinkedHashMap<>();
        shapes.put("bentArrow", ShapeTypes.BENT_ARROW);
        shapes.put("bentUpArrow", ShapeTypes.BENT_UP_ARROW);
        shapes.put("circularArrow", ShapeTypes.CIRCULAR_ARROW);
        shapes.put("curvedDownArrow", ShapeTypes.CURVED_DOWN_ARROW);
        shapes.put("curvedUpArrow", ShapeTypes.CURVED_UP_ARROW);
        shapes.put("leftCircularArrow", ShapeTypes.LEFT_CIRCULAR_ARROW);
        shapes.put("leftRightCircularArrow", ShapeTypes.LEFT_RIGHT_CIRCULAR_ARROW);
        shapes.put("leftRightUpArrow", ShapeTypes.LEFT_RIGHT_UP_ARROW);
        shapes.put("leftUpArrow", ShapeTypes.LEFT_UP_ARROW);
        shapes.put("notchedRightArrow", ShapeTypes.NOTCHED_RIGHT_ARROW);
        shapes.put("quadArrow", ShapeTypes.QUAD_ARROW);
        shapes.put("stripedRightArrow", ShapeTypes.STRIPED_RIGHT_ARROW);
        shapes.put("swooshArrow", ShapeTypes.SWOOSH_ARROW);
        shapes.put("upDownArrow", ShapeTypes.UP_DOWN_ARROW);
        shapes.put("uturnArrow", ShapeTypes.UTURN_ARROW);

        Path workbookPath = Path.of("target", "unsupported-arrows.xlsx");
        Path pdfPath = Path.of("target", "unsupported-arrows.pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            for (Map.Entry<String, Integer> entry : shapes.entrySet()) {
                int index = workbook.getNumberOfSheets();
                var sheet = workbook.createSheet(entry.getKey());
                sheet.createRow(0).createCell(0).setCellValue(" ");
                for (int c = 0; c < 5; c++) {
                    sheet.setColumnWidth(c, 12 * 256);
                }
                for (int r = 0; r < 5; r++) {
                    if (sheet.getRow(r) == null) {
                        sheet.createRow(r);
                    }
                    sheet.getRow(r).setHeightInPoints(20);
                }
                XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
                XSSFSimpleShape shape = drawing.createSimpleShape(
                        new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 4, 4));
                shape.setShapeType(entry.getValue());
                shape.setFillColor(0xFF, 0x99, 0xCC);
                shape.setLineStyleColor(0xFF, 0x99, 0xCC);
                shape.setLineWidth(0.5);
                workbook.setPrintArea(index, "A1:E5");
            }
            try (var out = Files.newOutputStream(workbookPath)) {
                workbook.write(out);
            }
        }

        ConversionResult result = Xlsx2Pdf.convert(workbookPath, pdfPath,
                ConversionOptions.builder()
                        .dpi(150)
                        .textLayerMode(TextLayerMode.NONE)
                        .build());
        assertThat(result.pageCount()).isEqualTo(shapes.size());
        assertThat(result.warnings()).isEmpty();

        try (PDDocument document = Loader.loadPDF(pdfPath.toFile())) {
            assertThat(document.getNumberOfPages()).isEqualTo(shapes.size());
            PDFRenderer renderer = new PDFRenderer(document);
            int page = 0;
            for (String name : shapes.keySet()) {
                BufferedImage image = renderer.renderImageWithDPI(page++, 150, ImageType.RGB);
                Ratio ratio = pinkRatio(image);
                assertThat(ratio.pixels)
                        .as("%s should draw filled pixels", name)
                        .isGreaterThan(100);
                assertThat(ratio.ratio)
                        .as("%s should not be a solid rectangle", name)
                        .isBetween(0.10, 0.93);
            }
        }
    }

    private static Ratio pinkRatio(BufferedImage image) {
        int minX = Integer.MAX_VALUE;
        int minY = Integer.MAX_VALUE;
        int maxX = -1;
        int maxY = -1;
        int pixels = 0;
        for (int y = 0; y < image.getHeight(); y++) {
            for (int x = 0; x < image.getWidth(); x++) {
                if (isPink(image.getRGB(x, y))) {
                    pixels++;
                    minX = Math.min(minX, x);
                    minY = Math.min(minY, y);
                    maxX = Math.max(maxX, x);
                    maxY = Math.max(maxY, y);
                }
            }
        }
        if (pixels == 0) {
            return new Ratio(0, 0);
        }
        double area = (maxX - minX + 1d) * (maxY - minY + 1d);
        return new Ratio(pixels, pixels / area);
    }

    private static boolean isPink(int rgb) {
        int red = (rgb >> 16) & 0xFF;
        int green = (rgb >> 8) & 0xFF;
        int blue = rgb & 0xFF;
        return red > 235 && green > 120 && green < 205 && blue > 175;
    }

    private record Ratio(int pixels, double ratio) {
    }
}
