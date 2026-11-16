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
 * Covers the basic-shape presets that were previously approximated by their
 * bounding rectangle.
 */
class BasicShapePresetsTest {

    @Test
    void allUnsupportedBasicPresetsRenderWithTheirPresetGeometry() throws Exception {
        Map<String, Integer> shapes = new LinkedHashMap<>();
        shapes.put("arc", ShapeTypes.ARC);
        shapes.put("bevel", ShapeTypes.BEVEL);
        shapes.put("blockArc", ShapeTypes.BLOCK_ARC);
        shapes.put("bracePair", ShapeTypes.BRACE_PAIR);
        shapes.put("bracketPair", ShapeTypes.BRACKET_PAIR);
        shapes.put("can", ShapeTypes.CAN);
        shapes.put("chartPlus", ShapeTypes.CHART_PLUS);
        shapes.put("chartStar", ShapeTypes.CHART_STAR);
        shapes.put("chartX", ShapeTypes.CHART_X);
        shapes.put("chevron", ShapeTypes.CHEVRON);
        shapes.put("chord", ShapeTypes.CHORD);
        shapes.put("cloud", ShapeTypes.CLOUD);
        shapes.put("corner", ShapeTypes.CORNER);
        shapes.put("cornerTabs", ShapeTypes.CORNER_TABS);
        shapes.put("cube", ShapeTypes.CUBE);
        shapes.put("decagon", ShapeTypes.DECAGON);
        shapes.put("diagStripe", ShapeTypes.DIAG_STRIPE);
        shapes.put("diamond", ShapeTypes.DIAMOND);
        shapes.put("dodecagon", ShapeTypes.DODECAGON);
        shapes.put("donut", ShapeTypes.DONUT);
        shapes.put("doubleWave", ShapeTypes.DOUBLE_WAVE);
        shapes.put("frame", ShapeTypes.FRAME);
        shapes.put("funnel", ShapeTypes.FUNNEL);
        shapes.put("gear6", ShapeTypes.GEAR_6);
        shapes.put("gear9", ShapeTypes.GEAR_9);
        shapes.put("halfFrame", ShapeTypes.HALF_FRAME);
        shapes.put("heart", ShapeTypes.HEART);
        shapes.put("heptagon", ShapeTypes.HEPTAGON);
        shapes.put("hexagon", ShapeTypes.HEXAGON);
        shapes.put("leftBrace", ShapeTypes.LEFT_BRACE);
        shapes.put("leftBracket", ShapeTypes.LEFT_BRACKET);
        shapes.put("lightningBolt", ShapeTypes.LIGHTNING_BOLT);
        shapes.put("lineInv", ShapeTypes.LINE_INV);
        shapes.put("mathDivide", ShapeTypes.MATH_DIVIDE);
        shapes.put("mathEqual", ShapeTypes.MATH_EQUAL);
        shapes.put("mathMinus", ShapeTypes.MATH_MINUS);
        shapes.put("mathMultiply", ShapeTypes.MATH_MULTIPLY);
        shapes.put("mathNotEqual", ShapeTypes.MATH_NOT_EQUAL);
        shapes.put("mathPlus", ShapeTypes.MATH_PLUS);
        shapes.put("moon", ShapeTypes.MOON);
        shapes.put("nonIsoscelesTrapezoid", ShapeTypes.NON_ISOSCELES_TRAPEZOID);
        shapes.put("noSmoking", ShapeTypes.NO_SMOKING);
        shapes.put("octagon", ShapeTypes.OCTAGON);
        shapes.put("parallelogram", ShapeTypes.PARALLELOGRAM);
        shapes.put("pentagon", ShapeTypes.PENTAGON);
        shapes.put("pie", ShapeTypes.PIE);
        shapes.put("pieWedge", ShapeTypes.PIE_WEDGE);
        shapes.put("plaque", ShapeTypes.PLAQUE);
        shapes.put("plaqueTabs", ShapeTypes.PLAQUE_TABS);
        shapes.put("plus", ShapeTypes.PLUS);
        shapes.put("rightBrace", ShapeTypes.RIGHT_BRACE);
        shapes.put("rightBracket", ShapeTypes.RIGHT_BRACKET);
        shapes.put("rtTriangle", ShapeTypes.RT_TRIANGLE);
        shapes.put("smileyFace", ShapeTypes.SMILEY_FACE);
        shapes.put("squareTabs", ShapeTypes.SQUARE_TABS);
        shapes.put("sun", ShapeTypes.SUN);
        shapes.put("teardrop", ShapeTypes.TEARDROP);
        shapes.put("trapezoid", ShapeTypes.TRAPEZOID);
        shapes.put("triangle", ShapeTypes.TRIANGLE);
        shapes.put("wave", ShapeTypes.WAVE);

        Path workbookPath = Path.of("target", "basic-presets.xlsx");
        Path pdfPath = Path.of("target", "basic-presets.pdf");
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
                        new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 5, 5));
                shape.setShapeType(entry.getValue());
                shape.setFillColor(0xFF, 0x99, 0xCC);
                shape.setLineStyleColor(0xFF, 0x99, 0xCC);
                shape.setLineWidth(0.5);
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
            assertThat(document.getNumberOfPages()).isEqualTo(shapes.size());
            PDFRenderer renderer = new PDFRenderer(document);
            int page = 0;
            for (String name : shapes.keySet()) {
                BufferedImage image = renderer.renderImageWithDPI(page++, 150, ImageType.RGB);
                assertThat(pinkPixels(image))
                        .as("%s should draw filled pixels", name)
                        .isGreaterThan(100);
            }
        }
    }

    private static int pinkPixels(BufferedImage image) {
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
