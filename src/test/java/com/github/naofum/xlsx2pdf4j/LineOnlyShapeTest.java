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
 * Arc and brace/bracket presets are outline-only shapes. They must not be
 * filled as a solid area (which would draw radial lines for an arc).
 */
class LineOnlyShapeTest {

    @Test
    void arcAndBraceBracketPresetsDrawLinesWithoutFill() throws Exception {
        Map<String, Integer> shapes = new LinkedHashMap<>();
        shapes.put("arc", ShapeTypes.ARC);
        shapes.put("bracePair", ShapeTypes.BRACE_PAIR);
        shapes.put("bracketPair", ShapeTypes.BRACKET_PAIR);
        shapes.put("leftBrace", ShapeTypes.LEFT_BRACE);
        shapes.put("rightBrace", ShapeTypes.RIGHT_BRACE);
        shapes.put("leftBracket", ShapeTypes.LEFT_BRACKET);
        shapes.put("rightBracket", ShapeTypes.RIGHT_BRACKET);

        Path workbookPath = Path.of("target", "line-only-shapes.xlsx");
        Path pdfPath = Path.of("target", "line-only-shapes.pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            for (Map.Entry<String, Integer> entry : shapes.entrySet()) {
                int index = workbook.getNumberOfSheets();
                var sheet = workbook.createSheet(entry.getKey());
                sheet.createRow(0).createCell(0).setCellValue(" ");
                for (int c = 0; c < 6; c++) {
                    sheet.setColumnWidth(c, 12 * 256);
                }
                for (int r = 0; r < 6; r++) {
                    if (sheet.getRow(r) == null) {
                        sheet.createRow(r);
                    }
                    sheet.getRow(r).setHeightInPoints(20);
                }
                XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
                XSSFSimpleShape shape = drawing.createSimpleShape(
                        new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 5, 5));
                shape.setShapeType(entry.getValue());
                shape.setFillColor(0xFF, 0x99, 0xCC);
                shape.setLineStyleColor(0, 0, 0);
                shape.setLineWidth(1.0);
                workbook.setPrintArea(index, "A1:F6");
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
            PDFRenderer renderer = new PDFRenderer(document);
            int page = 0;
            for (String name : shapes.keySet()) {
                BufferedImage image = renderer.renderImageWithDPI(page++, 150, ImageType.RGB);
                Counts counts = counts(image);
                assertThat(counts.pink)
                        .as("%s must not be filled", name)
                        .isLessThan(20);
                assertThat(counts.dark)
                        .as("%s should draw an outline", name)
                        .isGreaterThan(30);
            }
        }
    }

    private static Counts counts(BufferedImage image) {
        int pink = 0;
        int dark = 0;
        for (int y = 0; y < image.getHeight(); y++) {
            for (int x = 0; x < image.getWidth(); x++) {
                int rgb = image.getRGB(x, y);
                int red = (rgb >> 16) & 0xFF;
                int green = (rgb >> 8) & 0xFF;
                int blue = rgb & 0xFF;
                if (red > 235 && green > 120 && green < 205 && blue > 175) {
                    pink++;
                }
                if (red < 100 && green < 100 && blue < 100) {
                    dark++;
                }
            }
        }
        return new Counts(pink, dark);
    }

    private record Counts(int pink, int dark) {
    }
}
