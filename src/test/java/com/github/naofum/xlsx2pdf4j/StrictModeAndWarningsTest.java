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
import com.github.naofum.xlsx2pdf4j.api.ConversionWarning;
import com.github.naofum.xlsx2pdf4j.api.ErrorCode;
import com.github.naofum.xlsx2pdf4j.api.Xlsx2Pdf;
import com.github.naofum.xlsx2pdf4j.api.Xlsx2PdfException;
import com.github.naofum.xlsx2pdf4j.api.WarningCode;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class StrictModeAndWarningsTest {

    @Test
    void approximatedShapeProducesWarningButStrictModeContinues() throws Exception {
        Path workbookPath = Path.of("target", "warning-approximated.xlsx");
        Path normalPdf = Path.of("target", "warning-approximated.pdf");
        Path strictPdf = Path.of("target", "warning-approximated-strict.pdf");
        Files.deleteIfExists(normalPdf);
        Files.deleteIfExists(strictPdf);
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Shapes");
            sheet.createRow(0).createCell(0).setCellValue("x");
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            XSSFSimpleShape approximated = drawing.createSimpleShape(
                    new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 5, 5));
            approximated.setShapeType(ShapeTypes.ACTION_BUTTON_BLANK);
            approximated.setFillColor(0xC0, 0x00, 0x00);
            workbook.setPrintArea(0, "A1:F8");
            write(workbook, workbookPath);
        }

        ConversionResult normal = Xlsx2Pdf.convert(workbookPath, normalPdf,
                ConversionOptions.builder().dpi(100).build());
        assertThat(normal.warnings()).anySatisfy(warning -> {
            assertThat(warning.code()).isEqualTo(WarningCode.SHAPE_APPROXIMATED);
            assertThat(warning.sheetName()).isEqualTo("Shapes");
            assertThat(warning.location()).isNotBlank();
            assertThat(warning.message()).contains("approximated");
        });

        ConversionResult strict = Xlsx2Pdf.convert(workbookPath, strictPdf,
                ConversionOptions.builder().dpi(100).strict(true).build());
        assertThat(strict.warnings())
                .extracting(ConversionWarning::code)
                .contains(WarningCode.SHAPE_APPROXIMATED);
        assertThat(strictPdf).exists();
    }

    @Test
    void unsupportedImageIsCriticalInStrictModeAndLeavesNoOutput() throws Exception {
        Path workbookPath = Path.of("target", "warning-image.xlsx");
        Path normalPdf = Path.of("target", "warning-image.pdf");
        Path strictPdf = Path.of("target", "warning-image-strict.pdf");
        Files.deleteIfExists(normalPdf);
        Files.deleteIfExists(strictPdf);
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Images");
            sheet.createRow(0).createCell(0).setCellValue("x");
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            int pictureIndex = workbook.addPicture(
                    new byte[]{1, 2, 3, 4, 5}, Workbook.PICTURE_TYPE_EMF);
            XSSFPicture picture = drawing.createPicture(
                    new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 5, 5), pictureIndex);
            picture.getCTPicture().getNvPicPr().getCNvPr().setName("fake-emf");
            workbook.setPrintArea(0, "A1:F8");
            write(workbook, workbookPath);
        }

        ConversionResult normal = Xlsx2Pdf.convert(workbookPath, normalPdf,
                ConversionOptions.builder().dpi(100).build());
        assertThat(normal.warnings()).anySatisfy(warning -> {
            assertThat(warning.code()).isEqualTo(WarningCode.IMAGE_UNSUPPORTED);
            assertThat(warning.sheetName()).isEqualTo("Images");
            assertThat(warning.location()).isEqualTo("fake-emf");
        });

        assertThatThrownBy(() -> Xlsx2Pdf.convert(workbookPath, strictPdf,
                ConversionOptions.builder().dpi(100).strict(true).build()))
                .isInstanceOfSatisfying(Xlsx2PdfException.class, failure -> {
                    assertThat(failure.code()).isEqualTo(ErrorCode.UNSUPPORTED);
                    assertThat(failure.sheetName()).isEqualTo("Images");
                    assertThat(failure.location()).isEqualTo("fake-emf");
                });
        assertThat(strictPdf).doesNotExist();
    }

    @Test
    void emptySheetProducesPageLimitWarningButNotStrictFailure() throws Exception {
        Path workbookPath = Path.of("target", "warning-empty-sheet.xlsx");
        Path normalPdf = Path.of("target", "warning-empty-sheet.pdf");
        Path strictPdf = Path.of("target", "warning-empty-sheet-strict.pdf");
        Files.deleteIfExists(normalPdf);
        Files.deleteIfExists(strictPdf);
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet content = workbook.createSheet("Content");
            content.createRow(0).createCell(0).setCellValue("x");
            workbook.createSheet("Empty");
            workbook.setPrintArea(0, "A1:B2");
            write(workbook, workbookPath);
        }

        ConversionResult normal = Xlsx2Pdf.convert(workbookPath, normalPdf,
                ConversionOptions.builder().dpi(100).build());
        assertThat(normal.warnings()).anySatisfy(warning -> {
            assertThat(warning.code()).isEqualTo(WarningCode.PAGE_LIMIT);
            assertThat(warning.sheetName()).isEqualTo("Empty");
            assertThat(warning.message()).contains("no printable content");
        });
        assertThat(normal.pagesBySheet()).containsOnlyKeys("Content");

        ConversionResult strict = Xlsx2Pdf.convert(workbookPath, strictPdf,
                ConversionOptions.builder().dpi(100).strict(true).build());
        assertThat(strict.warnings())
                .extracting(ConversionWarning::code)
                .contains(WarningCode.PAGE_LIMIT);
        assertThat(strictPdf).exists();
    }

    @Test
    void uncachedFormulaProducesWarningAndStrictFailureWithoutOutput() throws Exception {
        Path workbookPath = Path.of("target", "warning-formula.xlsx");
        Path normalPdf = Path.of("target", "warning-formula.pdf");
        Path strictPdf = Path.of("target", "warning-formula-strict.pdf");
        Files.deleteIfExists(normalPdf);
        Files.deleteIfExists(strictPdf);
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Formula");
            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue("x");
            row.createCell(1).setCellFormula("FOO(1)");
            row.createCell(2).setCellFormula("1+1");
            workbook.setPrintArea(0, "A1:C3");
            write(workbook, workbookPath);
        }

        ConversionResult normal = Xlsx2Pdf.convert(workbookPath, normalPdf,
                ConversionOptions.builder().dpi(100).build());
        assertThat(normal.warnings()).anySatisfy(warning -> {
            assertThat(warning.code()).isEqualTo(WarningCode.FORMULA_NO_CACHE);
            assertThat(warning.sheetName()).isEqualTo("Formula");
            assertThat(warning.location()).isEqualTo("B1");
            assertThat(warning.message()).contains("not cached");
        });
        assertThat(normal.warnings())
                .filteredOn(warning -> warning.code() == WarningCode.FORMULA_NO_CACHE)
                .extracting(ConversionWarning::location)
                .containsExactly("B1");

        assertThatThrownBy(() -> Xlsx2Pdf.convert(workbookPath, strictPdf,
                ConversionOptions.builder().dpi(100).strict(true).build()))
                .isInstanceOfSatisfying(Xlsx2PdfException.class, failure -> {
                    assertThat(failure.code()).isEqualTo(ErrorCode.UNSUPPORTED);
                    assertThat(failure.sheetName()).isEqualTo("Formula");
                    assertThat(failure.location()).isEqualTo("B1");
                });
        assertThat(strictPdf).doesNotExist();
    }

    @Test
    void chartProducesSkippedWarningButStrictModeContinues() throws Exception {
        Path workbookPath = Path.of("target", "warning-chart.xlsx");
        Path normalPdf = Path.of("target", "warning-chart.pdf");
        Path strictPdf = Path.of("target", "warning-chart-strict.pdf");
        Files.deleteIfExists(normalPdf);
        Files.deleteIfExists(strictPdf);
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Chart");
            sheet.createRow(0).createCell(0).setCellValue(1);
            sheet.createRow(1).createCell(0).setCellValue(2);
            XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
            drawing.createChart(new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 6, 10));
            workbook.setPrintArea(0, "A1:H12");
            write(workbook, workbookPath);
        }

        ConversionResult normal = Xlsx2Pdf.convert(workbookPath, normalPdf,
                ConversionOptions.builder().dpi(100).build());
        assertThat(normal.warnings()).anySatisfy(warning -> {
            assertThat(warning.code()).isEqualTo(WarningCode.CHART_SKIPPED);
            assertThat(warning.sheetName()).isEqualTo("Chart");
            assertThat(warning.location()).isNotBlank();
            assertThat(warning.message()).contains("chart");
        });

        ConversionResult strict = Xlsx2Pdf.convert(workbookPath, strictPdf,
                ConversionOptions.builder().dpi(100).strict(true).build());
        assertThat(strict.warnings())
                .extracting(ConversionWarning::code)
                .contains(WarningCode.CHART_SKIPPED);
        assertThat(strictPdf).exists();
    }

    @Test
    void warningNormalizesMissingContext() {
        ConversionWarning warning = new ConversionWarning(
                WarningCode.TEXT_CLIPPED, null, null, null);
        assertThat(warning.sheetName()).isEmpty();
        assertThat(warning.location()).isEmpty();
        assertThat(warning.message()).isEmpty();
        assertThatThrownBy(() -> new ConversionWarning(null, "Sheet", "A1", "message"))
                .isInstanceOf(NullPointerException.class);
    }

    private static void write(XSSFWorkbook workbook, Path path) throws Exception {
        try (OutputStream out = Files.newOutputStream(path)) {
            workbook.write(out);
        }
    }
}
