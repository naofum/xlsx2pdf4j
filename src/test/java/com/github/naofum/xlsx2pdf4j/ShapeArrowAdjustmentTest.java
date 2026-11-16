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

import java.awt.image.BufferedImage;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

class ShapeArrowAdjustmentTest {

    @Test
    void leftRightArrowUsesAdj1ForShaftWidth() throws Exception {
        Path workbookPath = Path.of("target", "shape-arrow-adjustment.xlsx");
        Path pdfPath = Path.of("target", "shape-arrow-adjustment.pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Arrow");
            sheet.createRow(0).createCell(0).setCellValue(" ");
            for (int c = 0; c < 8; c++) {
                sheet.setColumnWidth(c, 10 * 256);
            }
            for (int r = 1; r <= 2; r++) {
                sheet.createRow(r).setHeightInPoints(20);
            }
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            XSSFSimpleShape shape = drawing.createSimpleShape(
                    new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 8, 3));
            shape.setShapeType(ShapeTypes.LEFT_RIGHT_ARROW);
            shape.setFillColor(0xFF, 0x99, 0xCC);
            shape.setLineStyleColor(0, 0, 0);
            shape.setLineWidth(0.75);
            setAdjustments(shape);

            workbook.setPrintArea(0, "A1:I5");
            try (OutputStream out = Files.newOutputStream(workbookPath)) {
                workbook.write(out);
            }
        }

        BufferedImage page = render(workbookPath, pdfPath);
        int[] bounds = pinkBounds(page);
        int height = bounds[3] - bounds[1] + 1;
        int x = (bounds[0] + bounds[2]) / 2;
        int thickness = verticalPinkThickness(page, x, bounds);
        double ratio = thickness / (double) height;
        assertThat(ratio)
                .as("adj1=54870 should make the shaft about 55%% of the height")
                .isBetween(0.45, 0.65);
    }

    @Test
    void cardinalArrowsUseAdj1ForShaftWidth() throws Exception {
        assertArrowShaft(ShapeTypes.RIGHT_ARROW, "right", true);
        assertArrowShaft(ShapeTypes.LEFT_ARROW, "left", true);
        assertArrowShaft(ShapeTypes.DOWN_ARROW, "down", false);
        assertArrowShaft(ShapeTypes.UP_ARROW, "up", false);
    }

    private void assertArrowShaft(int shapeType, String name, boolean horizontal) throws Exception {
        Path workbookPath = Path.of("target", "shape-arrow-" + name + ".xlsx");
        Path pdfPath = Path.of("target", "shape-arrow-" + name + ".pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Arrow");
            sheet.createRow(0).createCell(0).setCellValue(" ");
            for (int c = 0; c < 9; c++) {
                sheet.setColumnWidth(c, 10 * 256);
            }
            for (int r = 1; r <= 11; r++) {
                sheet.createRow(r).setHeightInPoints(20);
            }
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = horizontal
                    ? new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 8, 3)
                    : new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 3, 11);
            XSSFSimpleShape shape = drawing.createSimpleShape(anchor);
            shape.setShapeType(shapeType);
            shape.setFillColor(0xFF, 0x99, 0xCC);
            shape.setLineStyleColor(0, 0, 0);
            shape.setLineWidth(0.75);
            setAdjustments(shape);
            workbook.setPrintArea(0, "A1:I12");
            try (OutputStream out = Files.newOutputStream(workbookPath)) {
                workbook.write(out);
            }
        }

        BufferedImage page = render(workbookPath, pdfPath);
        int[] bounds = pinkBounds(page);
        double ratio;
        if (horizontal) {
            int height = bounds[3] - bounds[1] + 1;
            int x = (bounds[0] + bounds[2]) / 2;
            ratio = verticalPinkThickness(page, x, bounds) / (double) height;
        } else {
            int width = bounds[2] - bounds[0] + 1;
            int y = (bounds[1] + bounds[3]) / 2;
            ratio = horizontalPinkThickness(page, y, bounds) / (double) width;
        }
        assertThat(ratio)
                .as("%s arrow adj1=54870 shaft ratio", name)
                .isBetween(0.45, 0.65);
    }

    private static void setAdjustments(XSSFSimpleShape shape) {
        var guides = shape.getCTShape().getSpPr().getPrstGeom().getAvLst();
        var adj1 = guides.addNewGd();
        adj1.setName("adj1");
        adj1.setFmla("val 54870");
        var adj2 = guides.addNewGd();
        adj2.setName("adj2");
        adj2.setFmla("val 67211");
    }

    private static BufferedImage render(Path workbookPath, Path pdfPath) throws Exception {
        Xlsx2Pdf.convert(workbookPath, pdfPath,
                ConversionOptions.builder()
                        .dpi(150)
                        .textLayerMode(TextLayerMode.NONE)
                        .build());
        try (PDDocument document = Loader.loadPDF(pdfPath.toFile())) {
            assertThat(document.getNumberOfPages()).isEqualTo(1);
            return new PDFRenderer(document).renderImageWithDPI(0, 150, ImageType.RGB);
        }
    }

    private static int verticalPinkThickness(BufferedImage image, int x, int[] bounds) {
        int thickness = 0;
        for (int y = bounds[1]; y <= bounds[3]; y++) {
            if (isPink(image.getRGB(x, y))) {
                thickness++;
            }
        }
        return thickness;
    }

    private static int horizontalPinkThickness(BufferedImage image, int y, int[] bounds) {
        int thickness = 0;
        for (int x = bounds[0]; x <= bounds[2]; x++) {
            if (isPink(image.getRGB(x, y))) {
                thickness++;
            }
        }
        return thickness;
    }

    private static int[] pinkBounds(BufferedImage image) {
        int minX = Integer.MAX_VALUE;
        int minY = Integer.MAX_VALUE;
        int maxX = Integer.MIN_VALUE;
        int maxY = Integer.MIN_VALUE;
        for (int y = 0; y < image.getHeight(); y++) {
            for (int x = 0; x < image.getWidth(); x++) {
                if (isPink(image.getRGB(x, y))) {
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

    private static boolean isPink(int rgb) {
        return ((rgb >> 16) & 0xFF) > 220
                && ((rgb >> 8) & 0xFF) > 110
                && ((rgb >> 8) & 0xFF) < 200
                && (rgb & 0xFF) > 170
                && (rgb & 0xFF) < 235;
    }
}
