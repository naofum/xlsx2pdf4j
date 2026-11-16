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

import java.awt.image.BufferedImage;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Ensures the curved up/down arrows and swoosh arrow honor their adjustment
 * guides instead of always using the preset defaults.
 */
class CurvedArrowAdjustmentTest {

    @Test
    void curvedUpArrowUsesAdjustments() throws Exception {
        int defaultPixels = renderPixels(ShapeTypes.CURVED_UP_ARROW, Map.of());
        int adjustedPixels = renderPixels(ShapeTypes.CURVED_UP_ARROW,
                Map.of("adj1", "50000", "adj2", "65000", "adj3", "40000"));
        assertThat(adjustedPixels)
                .as("adjusted curvedUpArrow should differ from the default")
                .isNotEqualTo(defaultPixels);
    }

    @Test
    void curvedDownArrowUsesAdjustments() throws Exception {
        int defaultPixels = renderPixels(ShapeTypes.CURVED_DOWN_ARROW, Map.of());
        int adjustedPixels = renderPixels(ShapeTypes.CURVED_DOWN_ARROW,
                Map.of("adj1", "50000", "adj2", "65000", "adj3", "40000"));
        assertThat(adjustedPixels)
                .as("adjusted curvedDownArrow should differ from the default")
                .isNotEqualTo(defaultPixels);
    }

    @Test
    void swooshArrowUsesAdjustments() throws Exception {
        int defaultPixels = renderPixels(ShapeTypes.SWOOSH_ARROW, Map.of());
        int adjustedPixels = renderPixels(ShapeTypes.SWOOSH_ARROW,
                Map.of("adj1", "60000", "adj2", "40000"));
        assertThat(adjustedPixels)
                .as("adjusted swooshArrow should differ from the default")
                .isNotEqualTo(defaultPixels);
        assertThat(renderPixels(ShapeTypes.SWOOSH_ARROW, Map.of("adj1", "60000")))
                .as("swooshArrow head height should respond to adj1")
                .isNotEqualTo(defaultPixels);
        assertThat(renderPixels(ShapeTypes.SWOOSH_ARROW, Map.of("adj2", "40000")))
                .as("swooshArrow head width should respond to adj2")
                .isNotEqualTo(defaultPixels);
    }

    private static int renderPixels(int shapeType, Map<String, String> adjustments)
            throws Exception {
        Path workbook = Path.of("target", "curved-adjustment-" + shapeType
                + "-" + adjustments.hashCode() + ".xlsx");
        Path pdf = Path.of("target", "curved-adjustment-" + shapeType
                + "-" + adjustments.hashCode() + ".pdf");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("Shape");
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
            shape.setShapeType(shapeType);
            shape.setFillColor(0xFF, 0x99, 0xCC);
            shape.setLineStyleColor(0, 0, 0);
            shape.setLineWidth(1.0);
            var guides = shape.getCTShape().getSpPr().getPrstGeom().getAvLst();
            for (Map.Entry<String, String> adjustment : adjustments.entrySet()) {
                var guide = guides.addNewGd();
                guide.setName(adjustment.getKey());
                guide.setFmla("val " + adjustment.getValue());
            }
            wb.setPrintArea(0, "A1:G7");
            try (var out = Files.newOutputStream(workbook)) {
                wb.write(out);
            }
        }
        Xlsx2Pdf.convert(workbook, pdf, ConversionOptions.builder()
                .dpi(150)
                .textLayerMode(TextLayerMode.NONE)
                .build());
        try (PDDocument document = Loader.loadPDF(pdf.toFile())) {
            BufferedImage image = new PDFRenderer(document)
                    .renderImageWithDPI(0, 150, ImageType.RGB);
            int pixels = 0;
            for (int y = 0; y < image.getHeight(); y++) {
                for (int x = 0; x < image.getWidth(); x++) {
                    int rgb = image.getRGB(x, y);
                    int red = (rgb >> 16) & 0xFF;
                    int green = (rgb >> 8) & 0xFF;
                    int blue = rgb & 0xFF;
                    if (red > 235 && green > 120 && green < 205 && blue > 175) {
                        pixels++;
                    }
                }
            }
            return pixels;
        }
    }
}
