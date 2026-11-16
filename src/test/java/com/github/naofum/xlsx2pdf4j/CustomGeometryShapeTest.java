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
import org.openxmlformats.schemas.drawingml.x2006.main.CTAdjPoint2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTCustomGeometry2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGeomGuide;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGeomGuideList;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPath2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPath2DArcTo;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPath2DCubicBezierTo;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.STPathFillMode;

import java.awt.Rectangle;
import java.awt.image.BufferedImage;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * DrawingML custom geometry is not created by Apache POI's high level API, so
 * these tests build it through the schema classes and verify the rendered page.
 */
class CustomGeometryShapeTest {

    private static final int ORANGE = 0xFF9900;
    private static final int RED = 0xC00000;
    private static final int BLUE = 0x4472C4;
    private static final int GREEN = 0x70AD47;

    @Test
    void drawsCustomPathsCurvesArcFillModesAndGuides() throws Exception {
        Path workbookPath = Path.of("target", "custom-geometry.xlsx");
        Path pdfPath = Path.of("target", "custom-geometry.pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Custom");
            for (int c = 0; c < 10; c++) {
                sheet.setColumnWidth(c, 8 * 256);
            }
            for (int r = 0; r < 18; r++) {
                sheet.createRow(r);
            }

            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            addLShape(drawing, 1, 1, 4, 7, ORANGE);
            addArcShape(drawing, 6, 1, 9, 7, RED);
            addMultiPathShape(drawing, 1, 9, 4, 15, BLUE);
            addCurveShape(drawing, 6, 9, 9, 15, GREEN);

            workbook.setPrintArea(0, "A1:J18");
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

        Rectangle l = colorBounds(page, ORANGE, 28);
        assertThat(l.width).isGreaterThan(100);
        assertThat(l.height).isGreaterThan(50);
        assertThat(near(sample(page, l, 0.15, 0.15), ORANGE, 28))
                .as("L-shape left arm is filled")
                .isTrue();
        assertThat(near(sample(page, l, 0.85, 0.85), ORANGE, 28))
                .as("L-shape bottom arm is filled")
                .isTrue();
        assertThat(near(sample(page, l, 0.85, 0.15), 0xFFFFFF, 16))
                .as("L-shape cut-out stays white")
                .isTrue();

        Rectangle arc = colorBounds(page, RED, 28);
        assertThat(near(sample(page, arc, 0.50, 0.50), RED, 28))
                .as("arcTo dome is filled through its centre")
                .isTrue();
        assertThat(near(sample(page, arc, 0.05, 0.05), 0xFFFFFF, 16))
                .as("arcTo uses the clockwise (upper) semicircle")
                .isTrue();

        Rectangle multi = colorBounds(page, BLUE, 28);
        int outer = sample(page, multi, 0.10, 0.10);
        int darkened = sample(page, multi, 0.30, 0.30);
        int lightened = sample(page, multi, 0.50, 0.50);
        assertThat(luminance(darkened))
                .as("fill=darken path")
                .isLessThan(luminance(outer) - 10);
        assertThat(luminance(lightened))
                .as("fill=lighten path")
                .isGreaterThan(luminance(outer) + 10);

        Rectangle curve = colorBounds(page, GREEN, 28);
        assertThat(near(sample(page, curve, 0.90, 0.90), GREEN, 28))
                .as("cubic/quadratic curve body is filled")
                .isTrue();
        assertThat(near(sample(page, curve, 0.10, 0.10), 0xFFFFFF, 16))
                .as("area above the cubic curve stays white")
                .isTrue();
    }

    private static void addLShape(XSSFDrawing drawing, int c1, int r1, int c2, int r2,
                                  int rgb) {
        CTCustomGeometry2D geom = addShape(drawing, c1, r1, c2, r2, rgb);
        CTGeomGuideList guides = geom.getGdLst();
        guide(guides, "x1", "*/ w 2 5");
        guide(guides, "y1", "*/ h 3 5");

        CTPath2D path = addPath(geom);
        moveTo(path, "0", "0");
        lineTo(path, "x1", "0");
        lineTo(path, "x1", "y1");
        lineTo(path, "w", "y1");
        lineTo(path, "w", "h");
        lineTo(path, "0", "h");
        path.addNewClose();
    }

    private static void addArcShape(XSSFDrawing drawing, int c1, int r1, int c2, int r2,
                                    int rgb) {
        CTCustomGeometry2D geom = addShape(drawing, c1, r1, c2, r2, rgb);
        CTPath2D path = addPath(geom);
        moveTo(path, "0", "50");
        CTPath2DArcTo arc = path.addNewArcTo();
        arc.setWR(50L);
        arc.setHR(50L);
        arc.setStAng(10800000L);
        arc.setSwAng(10800000L);
        lineTo(path, "100", "100");
        lineTo(path, "0", "100");
        path.addNewClose();
    }

    private static void addMultiPathShape(XSSFDrawing drawing, int c1, int r1, int c2, int r2,
                                          int rgb) {
        CTCustomGeometry2D geom = addShape(drawing, c1, r1, c2, r2, rgb);

        CTPath2D outer = addPath(geom);
        moveTo(outer, "0", "0");
        lineTo(outer, "100", "0");
        lineTo(outer, "100", "100");
        lineTo(outer, "0", "100");
        outer.addNewClose();

        CTPath2D dark = addPath(geom);
        dark.setFill(STPathFillMode.DARKEN);
        dark.setStroke(false);
        moveTo(dark, "20", "20");
        lineTo(dark, "70", "20");
        lineTo(dark, "70", "70");
        lineTo(dark, "20", "70");
        dark.addNewClose();

        CTPath2D light = addPath(geom);
        light.setFill(STPathFillMode.LIGHTEN);
        light.setStroke(false);
        moveTo(light, "50", "35");
        lineTo(light, "65", "50");
        lineTo(light, "50", "65");
        lineTo(light, "35", "50");
        light.addNewClose();
    }

    private static void addCurveShape(XSSFDrawing drawing, int c1, int r1, int c2, int r2,
                                      int rgb) {
        CTCustomGeometry2D geom = addShape(drawing, c1, r1, c2, r2, rgb);

        CTPath2D path = addPath(geom);
        moveTo(path, "0", "100");
        CTPath2DCubicBezierTo cubic = path.addNewCubicBezTo();
        point(cubic.addNewPt(), "0", "35");
        point(cubic.addNewPt(), "35", "0");
        point(cubic.addNewPt(), "100", "0");
        lineTo(path, "100", "100");
        path.addNewClose();
    }

    private static CTCustomGeometry2D addShape(XSSFDrawing drawing, int c1, int r1,
                                               int c2, int r2, int rgb) {
        XSSFSimpleShape shape = drawing.createSimpleShape(
                new XSSFClientAnchor(0, 0, 0, 0, c1, r1, c2, r2));
        shape.setShapeType(ShapeTypes.RECT);
        shape.setFillColor((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
        shape.setLineStyleColor(0, 0, 0);
        shape.setLineWidth(1.25);

        CTShapeProperties properties = shape.getCTShape().getSpPr();
        properties.unsetPrstGeom();
        CTCustomGeometry2D geom = properties.addNewCustGeom();
        geom.addNewAvLst();
        geom.addNewGdLst();
        geom.addNewAhLst();
        geom.addNewCxnLst();
        var rect = geom.addNewRect();
        rect.setL(0L);
        rect.setT(0L);
        rect.setR(0L);
        rect.setB(0L);
        geom.addNewPathLst();
        return geom;
    }

    private static CTPath2D addPath(CTCustomGeometry2D geom) {
        CTPath2D path = geom.getPathLst().addNewPath();
        path.setW(100);
        path.setH(100);
        return path;
    }

    private static void guide(CTGeomGuideList guides, String name, String formula) {
        CTGeomGuide guide = guides.addNewGd();
        guide.setName(name);
        guide.setFmla(formula);
    }

    private static void moveTo(CTPath2D path, Object x, Object y) {
        point(path.addNewMoveTo().addNewPt(), x, y);
    }

    private static void lineTo(CTPath2D path, Object x, Object y) {
        point(path.addNewLnTo().addNewPt(), x, y);
    }

    private static void point(CTAdjPoint2D point, Object x, Object y) {
        point.setX(x);
        point.setY(y);
    }

    private static Rectangle colorBounds(BufferedImage image, int rgb, int tolerance) {
        int minX = Integer.MAX_VALUE;
        int minY = Integer.MAX_VALUE;
        int maxX = Integer.MIN_VALUE;
        int maxY = Integer.MIN_VALUE;
        for (int y = 0; y < image.getHeight(); y++) {
            for (int x = 0; x < image.getWidth(); x++) {
                if (near(image.getRGB(x, y), rgb, tolerance)) {
                    minX = Math.min(minX, x);
                    minY = Math.min(minY, y);
                    maxX = Math.max(maxX, x);
                    maxY = Math.max(maxY, y);
                }
            }
        }
        assertThat(maxX)
                .as("shape colour %06X must be present", rgb)
                .isGreaterThanOrEqualTo(minX);
        return new Rectangle(minX, minY, maxX - minX + 1, maxY - minY + 1);
    }

    private static int sample(BufferedImage image, Rectangle bounds, double fx, double fy) {
        int x = bounds.x + (int) Math.round(fx * (bounds.width - 1));
        int y = bounds.y + (int) Math.round(fy * (bounds.height - 1));
        return image.getRGB(x, y);
    }

    private static boolean near(int rgb, int expected, int tolerance) {
        return Math.abs(((rgb >> 16) & 0xFF) - ((expected >> 16) & 0xFF)) <= tolerance
                && Math.abs(((rgb >> 8) & 0xFF) - ((expected >> 8) & 0xFF)) <= tolerance
                && Math.abs((rgb & 0xFF) - (expected & 0xFF)) <= tolerance;
    }

    private static int luminance(int rgb) {
        return (((rgb >> 16) & 0xFF) + ((rgb >> 8) & 0xFF) + (rgb & 0xFF)) / 3;
    }
}
