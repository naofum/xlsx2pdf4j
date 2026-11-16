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
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.TextPosition;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.TextAlign;
import org.apache.poi.xssf.usermodel.TextAutofit;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextParagraph;
import org.apache.poi.xssf.usermodel.XSSFTextRun;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.awt.Rectangle;
import java.awt.image.BufferedImage;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class ShapeTextLayoutTest {

    @Test
    void paragraphAlignmentIsAppliedPerParagraph() throws Exception {
        Path workbookPath = Path.of("target", "shape-text-paragraphs.xlsx");
        Path pdfPath = Path.of("target", "shape-text-paragraphs.pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Text");
            sheet.createRow(0).createCell(0).setCellValue(" ");
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            XSSFSimpleShape shape = drawing.createSimpleShape(
                    new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 8, 8));
            shape.setShapeType(ShapeTypes.RECT);
            shape.setFillColor(0xEE, 0xEE, 0xEE);
            shape.setLineStyleColor(0xEE, 0xEE, 0xEE);
            shape.setWordWrap(false);
            shape.clearText();

            XSSFTextParagraph first = shape.addNewTextParagraph();
            addRun(first, "Left", 14);
            first.setTextAlign(TextAlign.LEFT);
            XSSFTextParagraph second = shape.addNewTextParagraph();
            addRun(second, "Center", 14);
            second.setTextAlign(TextAlign.CENTER);
            XSSFTextParagraph third = shape.addNewTextParagraph();
            addRun(third, "Right", 14);
            third.setTextAlign(TextAlign.RIGHT);

            workbook.setPrintArea(0, "A1:I10");
            write(workbook, workbookPath);
        }

        BufferedImage page = render(workbookPath, pdfPath);
        Rectangle shapeBounds = nonWhiteBounds(page);
        List<Rectangle> lines = darkLineBounds(page, shapeBounds);
        assertThat(lines).hasSize(3);

        assertThat((double) (lines.get(0).x - shapeBounds.x))
                .as("left paragraph starts near the left edge")
                .isLessThan(shapeBounds.width * 0.25);
        double secondCenter = lines.get(1).x + lines.get(1).width / 2d;
        assertThat(Math.abs(secondCenter - (shapeBounds.x + shapeBounds.width / 2d)))
                .as("center paragraph is centered")
                .isLessThan(shapeBounds.width * 0.10);
        assertThat((double) (shapeBounds.x + shapeBounds.width
                - (lines.get(2).x + lines.get(2).width)))
                .as("right paragraph ends near the right edge")
                .isLessThan(shapeBounds.width * 0.25);
    }

    @Test
    void anchorCtrCentersTheWholeParagraphBlock() throws Exception {
        Path workbookPath = Path.of("target", "shape-text-anchor-center.xlsx");
        Path pdfPath = Path.of("target", "shape-text-anchor-center.pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Text");
            sheet.createRow(0).createCell(0).setCellValue(" ");
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            XSSFSimpleShape shape = drawing.createSimpleShape(
                    new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 8, 8));
            shape.setShapeType(ShapeTypes.RECT);
            shape.setFillColor(0xEE, 0xEE, 0xEE);
            shape.setLineStyleColor(0xEE, 0xEE, 0xEE);
            shape.setWordWrap(false);
            shape.getCTShape().getTxBody().getBodyPr().setAnchorCtr(true);
            shape.clearText();

            XSSFTextParagraph first = shape.addNewTextParagraph();
            addRun(first, "WWWWWWWWWW", 14);
            first.setTextAlign(TextAlign.LEFT);
            XSSFTextParagraph second = shape.addNewTextParagraph();
            addRun(second, "Hi", 14);
            second.setTextAlign(TextAlign.LEFT);

            workbook.setPrintArea(0, "A1:I10");
            write(workbook, workbookPath);
        }

        BufferedImage page = render(workbookPath, pdfPath);
        Rectangle shapeBounds = nonWhiteBounds(page);
        List<Rectangle> lines = darkLineBounds(page, shapeBounds);
        assertThat(lines).hasSize(2);
        assertThat((double) lines.get(0).x)
                .as("the text block is centered as a whole")
                .isGreaterThan(shapeBounds.x + shapeBounds.width * 0.15);
        assertThat(Math.abs(lines.get(0).x - lines.get(1).x))
                .as("left-aligned paragraphs share the block's left edge")
                .isLessThan(3);
    }

    @Test
    void normAutofitShrinksTextToFitAndReducesTextLayerFontSize() throws Exception {
        Path workbookPath = Path.of("target", "shape-text-norm-autofit.xlsx");
        Path pdfPath = Path.of("target", "shape-text-norm-autofit.pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Text");
            sheet.createRow(0).createCell(0).setCellValue(" ");
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            XSSFSimpleShape shape = drawing.createSimpleShape(
                    new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 4, 4));
            shape.setShapeType(ShapeTypes.RECT);
            shape.setFillColor(0xEE, 0xEE, 0xEE);
            shape.setLineStyleColor(0xEE, 0xEE, 0xEE);
            shape.setWordWrap(true);
            shape.setTextAutofit(TextAutofit.NORMAL);
            shape.clearText();
            XSSFTextParagraph paragraph = shape.addNewTextParagraph();
            addRun(paragraph, "This long text must shrink to fit the small shape", 20);

            workbook.setPrintArea(0, "A1:H8");
            write(workbook, workbookPath);
        }

        List<Float> sizes = textLayerFontSizes(workbookPath, pdfPath);
        assertThat(sizes).isNotEmpty();
        assertThat(sizes.stream().max(Float::compare).orElse(0f))
                .as("normAutofit shrinks the effective font size")
                .isLessThan(18f);
    }

    @Test
    void normAutofitFontScaleIsApplied() throws Exception {
        Path workbookPath = Path.of("target", "shape-text-norm-scale.xlsx");
        Path pdfPath = Path.of("target", "shape-text-norm-scale.pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Text");
            sheet.createRow(0).createCell(0).setCellValue(" ");
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            XSSFSimpleShape shape = drawing.createSimpleShape(
                    new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 6, 6));
            shape.setShapeType(ShapeTypes.RECT);
            shape.setFillColor(0xEE, 0xEE, 0xEE);
            shape.setLineStyleColor(0xEE, 0xEE, 0xEE);
            shape.setTextAutofit(TextAutofit.NORMAL);
            shape.getCTShape().getTxBody().getBodyPr().getNormAutofit().setFontScale(50000);
            shape.clearText();
            XSSFTextParagraph paragraph = shape.addNewTextParagraph();
            addRun(paragraph, "Scaled text", 20);

            workbook.setPrintArea(0, "A1:H8");
            write(workbook, workbookPath);
        }

        List<Float> sizes = textLayerFontSizes(workbookPath, pdfPath);
        assertThat(sizes).isNotEmpty();
        assertThat(sizes.stream().max(Float::compare).orElse(0f))
                .as("normAutofit fontScale=50% halves the 20pt font")
                .isBetween(8f, 12f);
    }

    private static void addRun(XSSFTextParagraph paragraph, String text, double size) {
        XSSFTextRun run = paragraph.addNewTextRun();
        run.setText(text);
        run.setFontSize(size);
        run.setFontFamily("Arial", (byte) 0, (byte) 0, false);
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

    private static List<Float> textLayerFontSizes(Path workbookPath, Path pdfPath) throws Exception {
        Xlsx2Pdf.convert(workbookPath, pdfPath,
                ConversionOptions.builder()
                        .dpi(150)
                        .textLayerMode(TextLayerMode.INVISIBLE)
                        .build());
        List<Float> sizes = new ArrayList<>();
        try (PDDocument document = Loader.loadPDF(pdfPath.toFile())) {
            new PDFTextStripper() {
                @Override
                protected void writeString(String text, List<TextPosition> positions) {
                    if (text != null && text.length() > 3) {
                        for (TextPosition position : positions) {
                            sizes.add(position.getFontSizeInPt());
                        }
                    }
                }
            }.getText(document);
        }
        return sizes;
    }

    private static Rectangle nonWhiteBounds(BufferedImage image) {
        int minX = Integer.MAX_VALUE;
        int minY = Integer.MAX_VALUE;
        int maxX = Integer.MIN_VALUE;
        int maxY = Integer.MIN_VALUE;
        for (int y = 0; y < image.getHeight(); y++) {
            for (int x = 0; x < image.getWidth(); x++) {
                int rgb = image.getRGB(x, y);
                if (red(rgb) < 250 || green(rgb) < 250 || blue(rgb) < 250) {
                    minX = Math.min(minX, x);
                    minY = Math.min(minY, y);
                    maxX = Math.max(maxX, x);
                    maxY = Math.max(maxY, y);
                }
            }
        }
        assertThat(maxX).isGreaterThan(minX);
        assertThat(maxY).isGreaterThan(minY);
        return new Rectangle(minX, minY, maxX - minX + 1, maxY - minY + 1);
    }

    private static List<Rectangle> darkLineBounds(BufferedImage image, Rectangle area) {
        List<Rectangle> lines = new ArrayList<>();
        int start = -1;
        int end = -1;
        for (int y = area.y; y < area.y + area.height; y++) {
            boolean dark = false;
            for (int x = area.x; x < area.x + area.width; x++) {
                int rgb = image.getRGB(x, y);
                if (red(rgb) < 100 && green(rgb) < 100 && blue(rgb) < 100) {
                    dark = true;
                    break;
                }
            }
            if (dark) {
                if (start < 0) {
                    start = y;
                }
                end = y;
            } else if (start >= 0) {
                lines.add(lineBounds(image, area, start, end));
                start = -1;
            }
        }
        if (start >= 0) {
            lines.add(lineBounds(image, area, start, end));
        }
        return lines;
    }

    private static Rectangle lineBounds(BufferedImage image, Rectangle area, int y1, int y2) {
        int minX = Integer.MAX_VALUE;
        int maxX = Integer.MIN_VALUE;
        for (int y = y1; y <= y2; y++) {
            for (int x = area.x; x < area.x + area.width; x++) {
                int rgb = image.getRGB(x, y);
                if (red(rgb) < 100 && green(rgb) < 100 && blue(rgb) < 100) {
                    minX = Math.min(minX, x);
                    maxX = Math.max(maxX, x);
                }
            }
        }
        return new Rectangle(minX, y1, maxX - minX + 1, y2 - y1 + 1);
    }

    private static int red(int rgb) {
        return (rgb >> 16) & 0xFF;
    }

    private static int green(int rgb) {
        return (rgb >> 8) & 0xFF;
    }

    private static int blue(int rgb) {
        return rgb & 0xFF;
    }

    private static void write(XSSFWorkbook workbook, Path path) throws Exception {
        try (OutputStream out = Files.newOutputStream(path)) {
            workbook.write(out);
        }
    }
}
