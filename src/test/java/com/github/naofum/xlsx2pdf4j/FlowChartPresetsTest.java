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
 * Covers the flowchart presets that were previously approximated by their
 * bounding rectangle.
 */
class FlowChartPresetsTest {

    @Test
    void allUnsupportedFlowChartPresetsRenderWithTheirPresetGeometry() throws Exception {
        Map<String, Integer> shapes = new LinkedHashMap<>();
        shapes.put("collate", ShapeTypes.FLOW_CHART_COLLATE);
        shapes.put("delay", ShapeTypes.FLOW_CHART_DELAY);
        shapes.put("display", ShapeTypes.FLOW_CHART_DISPLAY);
        shapes.put("extract", ShapeTypes.FLOW_CHART_EXTRACT);
        shapes.put("inputOutput", ShapeTypes.FLOW_CHART_INPUT_OUTPUT);
        shapes.put("internalStorage", ShapeTypes.FLOW_CHART_INTERNAL_STORAGE);
        shapes.put("magneticDrum", ShapeTypes.FLOW_CHART_MAGNETIC_DRUM);
        shapes.put("magneticTape", ShapeTypes.FLOW_CHART_MAGNETIC_TAPE);
        shapes.put("manualInput", ShapeTypes.FLOW_CHART_MANUAL_INPUT);
        shapes.put("manualOperation", ShapeTypes.FLOW_CHART_MANUAL_OPERATION);
        shapes.put("merge", ShapeTypes.FLOW_CHART_MERGE);
        shapes.put("offlineStorage", ShapeTypes.FLOW_CHART_OFFLINE_STORAGE);
        shapes.put("onlineStorage", ShapeTypes.FLOW_CHART_ONLINE_STORAGE);
        shapes.put("or", ShapeTypes.FLOW_CHART_OR);
        shapes.put("preparation", ShapeTypes.FLOW_CHART_PREPARATION);
        shapes.put("punchedCard", ShapeTypes.FLOW_CHART_PUNCHED_CARD);
        shapes.put("punchedTape", ShapeTypes.FLOW_CHART_PUNCHED_TAPE);
        shapes.put("sort", ShapeTypes.FLOW_CHART_SORT);
        shapes.put("summingJunction", ShapeTypes.FLOW_CHART_SUMMING_JUNCTION);
        shapes.put("terminator", ShapeTypes.FLOW_CHART_TERMINATOR);

        Path workbookPath = Path.of("target", "flowchart-presets.xlsx");
        Path pdfPath = Path.of("target", "flowchart-presets.pdf");
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
                if (!"internalStorage".equals(name)) {
                    assertThat(ratio.ratio)
                            .as("%s should not be a solid rectangle", name)
                            .isBetween(0.10, 0.985);
                }
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
