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
import java.util.Arrays;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Verifies the ECMA-376 path structure of the preset shapes that were
 * previously approximated with unrelated geometry.
 */
class EcmaPresetGeometryTest {

    @Test
    void chartPresetsDrawLinesOverAFilledRectangle() throws Exception {
        BufferedImage plus = render(ShapeTypes.CHART_PLUS, Map.of());
        Bounds bounds = bounds(plus);
        assertThat(bounds.width()).isGreaterThan(200);
        assertThat(bounds.height()).isGreaterThan(80);

        assertDarkNear(plus, bounds.centerX(), bounds.centerY(), 1,
                "chartPlus line intersection");
        assertDarkNear(plus, bounds.centerX(), bounds.atY(0.25), 3,
                "chartPlus vertical line");
        assertDarkNear(plus, bounds.atX(0.25), bounds.centerY(), 3,
                "chartPlus horizontal line");
        assertPink(plus, bounds.atX(0.25), bounds.atY(0.25),
                "chartPlus should keep a filled background");
        assertPink(plus, bounds.atX(0.75), bounds.atY(0.75),
                "chartPlus should keep a filled background");

        BufferedImage star = render(ShapeTypes.CHART_STAR, Map.of());
        bounds = bounds(star);
        assertDarkNear(star, bounds.centerX(), bounds.centerY(), 1,
                "chartStar line intersection");
        assertDarkNear(star, bounds.atX(0.25), bounds.atY(0.25), 3,
                "chartStar first diagonal");
        assertDarkNear(star, bounds.atX(0.75), bounds.atY(0.25), 3,
                "chartStar second diagonal");
        assertDarkNear(star, bounds.centerX(), bounds.atY(0.25), 3,
                "chartStar vertical line");
        assertPink(star, bounds.atX(0.25), bounds.atY(0.75),
                "chartStar should keep a filled background");

        BufferedImage x = render(ShapeTypes.CHART_X, Map.of());
        bounds = bounds(x);
        assertDarkNear(x, bounds.centerX(), bounds.centerY(), 1,
                "chartX line intersection");
        assertDarkNear(x, bounds.atX(0.25), bounds.atY(0.25), 3,
                "chartX falling diagonal");
        assertDarkNear(x, bounds.atX(0.75), bounds.atY(0.25), 3,
                "chartX rising diagonal");
        assertPink(x, bounds.atX(0.25), bounds.atY(0.75),
                "chartX should keep a filled background");
    }

    @Test
    void lineInvDrawsOnlyTheFallingDiagonal() throws Exception {
        BufferedImage image = render(ShapeTypes.LINE_INV, Map.of());
        Bounds bounds = bounds(image);
        assertDarkNear(image, bounds.atX(0.25), bounds.atY(0.75), 4,
                "lineInv top-right to bottom-left diagonal");
        assertDarkNear(image, bounds.atX(0.75), bounds.atY(0.25), 4,
                "lineInv top-right to bottom-left diagonal");
        assertNoDarkNear(image, bounds.atX(0.25), bounds.atY(0.25), 6,
                "lineInv must not draw the other diagonal");
        assertNoDarkNear(image, bounds.atX(0.75), bounds.atY(0.75), 6,
                "lineInv must not draw the other diagonal");
    }

    @Test
    void bevelDrawsButtonShadingOnAllFourSides() throws Exception {
        BufferedImage image = render(ShapeTypes.BEVEL, Map.of());
        Bounds bounds = bounds(image);
        double band = Math.min(bounds.width(), bounds.height()) * 0.125 / 2d;
        int center = luminance(image, bounds.centerX(), bounds.centerY(), 1);
        int top = luminance(image, bounds.centerX(), (int) Math.round(bounds.top() + band), 1);
        int left = luminance(image, (int) Math.round(bounds.left() + band), bounds.centerY(), 1);
        int bottom = luminance(image, bounds.centerX(),
                (int) Math.round(bounds.bottom() - band), 1);
        int right = luminance(image, (int) Math.round(bounds.right() - band),
                bounds.centerY(), 1);

        assertThat(top).as("bevel top face").isGreaterThan(center + 5);
        assertThat(left).as("bevel left face").isGreaterThan(center + 5);
        assertThat(bottom).as("bevel bottom face").isLessThan(center - 5);
        assertThat(right).as("bevel right face").isLessThan(center - 5);
    }

    @Test
    void tabPresetsDrawOnlyFourCornerTabs() throws Exception {
        for (int shapeType : new int[]{ShapeTypes.CORNER_TABS,
                ShapeTypes.SQUARE_TABS, ShapeTypes.PLAQUE_TABS}) {
            BufferedImage image = render(shapeType, Map.of());
            Bounds bounds = bounds(image);
            double dx = Math.hypot(bounds.width(), bounds.height()) / 20d;
            int inset = Math.max(2, (int) Math.round(dx * 0.35));
            String name = "shape " + shapeType;

            assertWhite(image, bounds.centerX(), bounds.centerY(),
                    name + " center must stay empty");
            assertWhite(image, bounds.centerX(), bounds.atY(0.25),
                    name + " top middle must stay empty");
            assertWhite(image, bounds.atX(0.25), bounds.centerY(),
                    name + " left middle must stay empty");

            assertPink(image, bounds.left() + inset, bounds.top() + inset,
                    name + " top-left tab");
            assertPink(image, bounds.left() + inset, bounds.bottom() - inset,
                    name + " bottom-left tab");
            assertPink(image, bounds.right() - inset, bounds.top() + inset,
                    name + " top-right tab");
            assertPink(image, bounds.right() - inset, bounds.bottom() - inset,
                    name + " bottom-right tab");
        }
    }

    @Test
    void blockArcDefaultsToTheUpperHalf() throws Exception {
        BufferedImage image = render(ShapeTypes.BLOCK_ARC, Map.of());
        Bounds bounds = bounds(image);
        assertPink(image, bounds.centerX(), bounds.atY(0.10),
                "blockArc upper half should be filled");
        assertWhite(image, bounds.centerX(), bounds.atY(0.75),
                "blockArc lower half should stay empty");
    }

    @Test
    void canAndCubeShadeTheirFaces() throws Exception {
        BufferedImage can = render(ShapeTypes.CAN, Map.of());
        Bounds canBounds = bounds(can);
        double ss = Math.min(canBounds.width(), canBounds.height());
        int canTop = luminance(can, canBounds.centerX(),
                (int) Math.round(canBounds.top() + ss / 8d), 1);
        int canBody = luminance(can, canBounds.centerX(), canBounds.atY(0.65), 1);
        assertThat(canTop).as("can top ellipse should be lighter than the body")
                .isGreaterThan(canBody + 5);

        BufferedImage cube = render(ShapeTypes.CUBE, Map.of());
        Bounds cubeBounds = bounds(cube);
        ss = Math.min(cubeBounds.width(), cubeBounds.height());
        double y1 = ss * 0.25;
        int cubeTop = luminance(cube,
                (int) Math.round(cubeBounds.left() + cubeBounds.width() * 0.5),
                (int) Math.round(cubeBounds.top() + y1 * 0.5), 1);
        int cubeFront = luminance(cube,
                (int) Math.round(cubeBounds.left() + cubeBounds.width() * 0.35),
                cubeBounds.atY(0.65), 1);
        int cubeSide = luminance(cube,
                (int) Math.round(cubeBounds.right() - y1 * 0.5),
                cubeBounds.atY(0.60), 1);
        assertThat(cubeTop).as("cube top face should be lighter than the front")
                .isGreaterThan(cubeFront + 5);
        assertThat(cubeSide).as("cube side face should be darker than the front")
                .isLessThan(cubeFront - 5);
        assertDarkNear(cube,
                (int) Math.round(cubeBounds.left()
                        + (cubeBounds.width() - y1) / 2d),
                (int) Math.round(cubeBounds.top() + y1), 2,
                "cube front-top edge");
    }

    @Test
    void noSmokingDrawsAWideTopLeftToBottomRightBar() throws Exception {
        BufferedImage image = renderSquareNoSmoking();
        Bounds bounds = bounds(image);
        assertPink(image, bounds.centerX(), bounds.centerY(),
                "noSmoking bar center");
        assertPink(image, bounds.atX(0.35), bounds.atY(0.35),
                "noSmoking top-left half of the bar");
        assertPink(image, bounds.atX(0.65), bounds.atY(0.65),
                "noSmoking bottom-right half of the bar");
        assertWhite(image, bounds.atX(0.70), bounds.atY(0.30),
                "noSmoking upper-right opening");
        assertWhite(image, bounds.atX(0.30), bounds.atY(0.70),
                "noSmoking lower-left opening");
    }

    @Test
    void octagonUsesECMAFlatEdges() throws Exception {
        BufferedImage image = render(ShapeTypes.OCTAGON, Map.of());
        Bounds bounds = bounds(image);
        assertDarkNear(image, bounds.atX(0.25),
                (int) Math.round(bounds.top() + 1), 3,
                "octagon flat top edge");
    }

    @Test
    void plaqueUsesConcaveCorners() throws Exception {
        BufferedImage image = render(ShapeTypes.PLAQUE, Map.of());
        Bounds bounds = bounds(image);
        double ss = Math.min(bounds.width(), bounds.height());
        double x1 = ss * 16667d / 100000d;
        assertWhite(image,
                (int) Math.round(bounds.left() + x1 * 0.5),
                (int) Math.round(bounds.top() + x1 * 0.5),
                "plaque corner should be cut inward");
        assertPink(image,
                (int) Math.round(bounds.left() + x1 * 1.5),
                (int) Math.round(bounds.top() + x1 * 0.5),
                "plaque top edge beyond the corner arc");
    }

    @Test
    void nonIsoscelesTrapezoidHasDifferentTopAndBottomWidths() throws Exception {
        BufferedImage image = render(ShapeTypes.NON_ISOSCELES_TRAPEZOID, Map.of());
        Bounds bounds = bounds(image);
        assertWhite(image, bounds.atX(0.95), bounds.atY(0.05),
                "nonIsoscelesTrapezoid top-right corner should be cut");
        assertPink(image, bounds.atX(0.95), bounds.atY(0.95),
                "nonIsoscelesTrapezoid bottom edge should reach the right side");
    }

    @Test
    void presetsHonorTheirAdjustmentGuides() throws Exception {
        assertAdjustmentChanges(ShapeTypes.BEVEL, Map.of("adj", "40000"));
        assertAdjustmentChanges(ShapeTypes.PLAQUE, Map.of("adj", "40000"));
        assertAdjustmentChanges(ShapeTypes.TEARDROP, Map.of("adj", "180000"));
        assertAdjustmentChanges(ShapeTypes.DIAG_STRIPE, Map.of("adj", "80000"));
        assertAdjustmentChanges(ShapeTypes.HALF_FRAME, Map.of("adj1", "10000"));
        assertAdjustmentChanges(ShapeTypes.HALF_FRAME, Map.of("adj2", "60000"));
        assertAdjustmentChanges(ShapeTypes.DOUBLE_WAVE, Map.of("adj1", "12000"));
        assertAdjustmentChanges(ShapeTypes.DOUBLE_WAVE, Map.of("adj2", "6000"));
        assertAdjustmentChanges(ShapeTypes.PLUS, Map.of("adj", "40000"));
        assertAdjustmentChanges(ShapeTypes.WAVE, Map.of("adj1", "18000"));
        assertAdjustmentChanges(ShapeTypes.WAVE, Map.of("adj2", "8000"));
        assertAdjustmentChanges(ShapeTypes.MATH_MULTIPLY, Map.of("adj1", "40000"));
        assertAdjustmentChanges(ShapeTypes.OCTAGON, Map.of("adj", "40000"));
        assertAdjustmentChanges(ShapeTypes.NO_SMOKING, Map.of("adj", "30000"));
        assertAdjustmentChanges(ShapeTypes.BLOCK_ARC,
                Map.of("adj1", "5400000", "adj2", "0", "adj3", "30000"));
        assertAdjustmentChanges(ShapeTypes.CAN, Map.of("adj", "40000"));
        assertAdjustmentChanges(ShapeTypes.CUBE, Map.of("adj", "50000"));
        assertAdjustmentChanges(ShapeTypes.NON_ISOSCELES_TRAPEZOID,
                Map.of("adj1", "40000"));
        assertAdjustmentChanges(ShapeTypes.NON_ISOSCELES_TRAPEZOID,
                Map.of("adj2", "5000"));
    }

    private static void assertAdjustmentChanges(int shapeType,
                                                Map<String, String> adjustments)
            throws Exception {
        int defaultHash = pixelHash(render(shapeType, Map.of()));
        int adjustedHash = pixelHash(render(shapeType, adjustments));
        assertThat(adjustedHash)
                .as("%s should respond to %s", shapeType, adjustments)
                .isNotEqualTo(defaultHash);
    }

    private static BufferedImage render(int shapeType, Map<String, String> adjustments)
            throws Exception {
        Path workbook = Path.of("target", "ecma-preset-" + shapeType
                + "-" + adjustments.hashCode() + ".xlsx");
        Path pdf = Path.of("target", "ecma-preset-" + shapeType
                + "-" + adjustments.hashCode() + ".pdf");
        Files.createDirectories(workbook.getParent());
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("Shape");
            sheet.createRow(0).createCell(0).setCellValue(" ");
            for (int c = 0; c < 8; c++) {
                sheet.setColumnWidth(c, 12 * 256);
            }
            for (int r = 0; r < 8; r++) {
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
            wb.setPrintArea(0, "A1:H8");
            try (var out = Files.newOutputStream(workbook)) {
                wb.write(out);
            }
        }
        Xlsx2Pdf.convert(workbook, pdf, ConversionOptions.builder()
                .dpi(150)
                .textLayerMode(TextLayerMode.NONE)
                .build());
        try (PDDocument document = Loader.loadPDF(pdf.toFile())) {
            return new PDFRenderer(document)
                    .renderImageWithDPI(0, 150, ImageType.RGB);
        }
    }

    private static BufferedImage renderSquareNoSmoking() throws Exception {
        Path workbook = Path.of("target", "ecma-preset-noSmoking-square.xlsx");
        Path pdf = Path.of("target", "ecma-preset-noSmoking-square.pdf");
        Files.createDirectories(workbook.getParent());
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("Shape");
            sheet.createRow(0).createCell(0).setCellValue(" ");
            for (int c = 0; c < 5; c++) {
                sheet.setColumnWidth(c, 20 * 256);
            }
            for (int r = 0; r < 5; r++) {
                if (sheet.getRow(r) == null) {
                    sheet.createRow(r);
                }
                sheet.getRow(r).setHeightInPoints(109);
            }
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            XSSFSimpleShape shape = drawing.createSimpleShape(
                    new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 4, 4));
            shape.setShapeType(ShapeTypes.NO_SMOKING);
            shape.setFillColor(0xFF, 0x99, 0xCC);
            shape.setLineStyleColor(0, 0, 0);
            shape.setLineWidth(1.0);
            wb.setPrintArea(0, "A1:E5");
            try (var out = Files.newOutputStream(workbook)) {
                wb.write(out);
            }
        }
        Xlsx2Pdf.convert(workbook, pdf, ConversionOptions.builder()
                .dpi(150)
                .textLayerMode(TextLayerMode.NONE)
                .build());
        try (PDDocument document = Loader.loadPDF(pdf.toFile())) {
            return new PDFRenderer(document)
                    .renderImageWithDPI(0, 150, ImageType.RGB);
        }
    }

    private static Bounds bounds(BufferedImage image) {
        int left = image.getWidth();
        int top = image.getHeight();
        int right = -1;
        int bottom = -1;
        for (int y = 0; y < image.getHeight(); y++) {
            for (int x = 0; x < image.getWidth(); x++) {
                if (!isWhite(image.getRGB(x, y))) {
                    left = Math.min(left, x);
                    top = Math.min(top, y);
                    right = Math.max(right, x);
                    bottom = Math.max(bottom, y);
                }
            }
        }
        assertThat(right).as("rendered shape bounds").isGreaterThan(left);
        assertThat(bottom).as("rendered shape bounds").isGreaterThan(top);
        return new Bounds(left, top, right + 1, bottom + 1);
    }

    private static int pixelHash(BufferedImage image) {
        int[] pixels = image.getRGB(0, 0, image.getWidth(), image.getHeight(),
                null, 0, image.getWidth());
        return Arrays.hashCode(pixels);
    }

    private static void assertDarkNear(BufferedImage image, int x, int y, int radius,
                                       String description) {
        assertThat(hasColorNear(image, x, y, radius, ColorKind.DARK))
                .as(description + " should be drawn as a dark line")
                .isTrue();
    }

    private static void assertNoDarkNear(BufferedImage image, int x, int y, int radius,
                                         String description) {
        assertThat(hasColorNear(image, x, y, radius, ColorKind.DARK))
                .as(description)
                .isFalse();
    }

    private static void assertPink(BufferedImage image, int x, int y, String description) {
        assertThat(hasColorNear(image, x, y, 2, ColorKind.PINK))
                .as(description + " should be filled")
                .isTrue();
    }

    private static void assertWhite(BufferedImage image, int x, int y, String description) {
        assertThat(hasColorNear(image, x, y, 2, ColorKind.WHITE))
                .as(description)
                .isTrue();
    }

    private static boolean hasColorNear(BufferedImage image, int x, int y, int radius,
                                        ColorKind kind) {
        for (int dy = -radius; dy <= radius; dy++) {
            for (int dx = -radius; dx <= radius; dx++) {
                int px = x + dx;
                int py = y + dy;
                if (px < 0 || py < 0 || px >= image.getWidth() || py >= image.getHeight()) {
                    continue;
                }
                int rgb = image.getRGB(px, py);
                if (kind.matches(rgb)) {
                    return true;
                }
            }
        }
        return false;
    }

    private static int luminance(BufferedImage image, int x, int y, int radius) {
        int total = 0;
        int count = 0;
        for (int dy = -radius; dy <= radius; dy++) {
            for (int dx = -radius; dx <= radius; dx++) {
                int px = x + dx;
                int py = y + dy;
                if (px < 0 || py < 0 || px >= image.getWidth() || py >= image.getHeight()) {
                    continue;
                }
                total += rgbLuminance(image.getRGB(px, py));
                count++;
            }
        }
        return count == 0 ? 0 : total / count;
    }

    private static int rgbLuminance(int rgb) {
        return (((rgb >> 16) & 0xFF) + ((rgb >> 8) & 0xFF) + (rgb & 0xFF)) / 3;
    }

    private static boolean isWhite(int rgb) {
        int red = (rgb >> 16) & 0xFF;
        int green = (rgb >> 8) & 0xFF;
        int blue = rgb & 0xFF;
        return red > 248 && green > 248 && blue > 248;
    }

    private enum ColorKind {
        DARK {
            @Override
            boolean matches(int rgb) {
                return rgbLuminance(rgb) < 110;
            }
        },
        PINK {
            @Override
            boolean matches(int rgb) {
                int red = (rgb >> 16) & 0xFF;
                int green = (rgb >> 8) & 0xFF;
                int blue = rgb & 0xFF;
                return red > 200 && green > 70 && green < 235 && blue > 130;
            }
        },
        WHITE {
            @Override
            boolean matches(int rgb) {
                return isWhite(rgb);
            }
        };

        abstract boolean matches(int rgb);
    }

    private record Bounds(int left, int top, int right, int bottom) {
        int width() {
            return right - left;
        }

        int height() {
            return bottom - top;
        }

        int centerX() {
            return (left + right) / 2;
        }

        int centerY() {
            return (top + bottom) / 2;
        }

        int atX(double fraction) {
            return (int) Math.round(left + width() * fraction);
        }

        int atY(double fraction) {
            return (int) Math.round(top + height() * fraction);
        }
    }
}
