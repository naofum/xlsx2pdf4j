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
import com.github.naofum.xlsx2pdf4j.api.Docx2Pdf;
import com.github.naofum.xlsx2pdf4j.api.DocxConversionOptions;
import com.github.naofum.xlsx2pdf4j.api.Pptx2Pdf;
import com.github.naofum.xlsx2pdf4j.api.PptxConversionOptions;
import com.github.naofum.xlsx2pdf4j.api.Xlsx2Pdf;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDDocumentOutline;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDOutlineItem;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

class PdfBookmarksTest {

    @Test
    void excelSheetBookmarksAreOptInAndPointToEachSheetsFirstPage() throws Exception {
        Path workbook = Path.of("target", "pdf-bookmarks", "sheets.xlsx");
        Path withBookmarks = Path.of("target", "pdf-bookmarks", "sheets-with-bookmarks.pdf");
        Path withoutBookmarks = Path.of("target", "pdf-bookmarks", "sheets-without-bookmarks.pdf");
        writeWorkbook(workbook);

        Xlsx2Pdf.convert(workbook, withBookmarks,
                ConversionOptions.builder()
                        .dpi(72)
                        .sheetBookmarks(true)
                        .build());
        Xlsx2Pdf.convert(workbook, withoutBookmarks,
                ConversionOptions.builder().dpi(72).build());

        try (PDDocument document = Loader.loadPDF(withBookmarks.toFile())) {
            PDDocumentOutline outline = document.getDocumentCatalog().getDocumentOutline();
            assertThat(outline).isNotNull();
            PDOutlineItem first = outline.getFirstChild();
            assertThat(first.getTitle()).isEqualTo("First");
            assertThat(document.getPages().indexOf(first.findDestinationPage(document))).isZero();

            PDOutlineItem second = first.getNextSibling();
            assertThat(second).isNotNull();
            assertThat(second.getTitle()).isEqualTo("Second");
            assertThat(document.getPages().indexOf(second.findDestinationPage(document))).isEqualTo(1);
        }
        try (PDDocument document = Loader.loadPDF(withoutBookmarks.toFile())) {
            assertThat(document.getDocumentCatalog().getDocumentOutline()).isNull();
        }
    }

    @Test
    void powerpointSlideTitleBecomesABookmark() throws Exception {
        String expected;
        try (XMLSlideShow presentation = new XMLSlideShow(
                Files.newInputStream(Path.of("testdata", "validation", "sample1.pptx")))) {
            expected = presentation.getSlides().get(0).getTitle().strip();
        }
        assertThat(expected).isNotBlank();

        Path pdf = Path.of("target", "pdf-bookmarks", "powerpoint.pdf");
        Pptx2Pdf.convert(
                Path.of("testdata", "validation", "sample1.pptx"), pdf,
                PptxConversionOptions.builder()
                        .dpi(72)
                        .slideBookmarks(true)
                        .build());

        try (PDDocument document = Loader.loadPDF(pdf.toFile())) {
            PDDocumentOutline outline = document.getDocumentCatalog().getDocumentOutline();
            assertThat(outline).isNotNull();
            PDOutlineItem item = outline.getFirstChild();
            assertThat(item).isNotNull();
            assertThat(item.getTitle()).isEqualTo(expected);
            assertThat(document.getPages().indexOf(item.findDestinationPage(document))).isZero();
        }
    }

    @Test
    void wordHeadingsBecomeNestedBookmarks() throws Exception {
        Path pdf = Path.of("target", "pdf-bookmarks", "word.pdf");
        Docx2Pdf.convert(
                Path.of("testdata", "validation", "sample1.docx"), pdf,
                DocxConversionOptions.builder()
                        .dpi(72)
                        .headingBookmarks(true)
                        .build());

        try (PDDocument document = Loader.loadPDF(pdf.toFile())) {
            PDDocumentOutline outline = document.getDocumentCatalog().getDocumentOutline();
            assertThat(outline).isNotNull();
            PDOutlineItem heading1 = outline.getFirstChild();
            assertThat(heading1).isNotNull();
            assertThat(heading1.getTitle()).isEqualTo("サービスのコンセプト");
            assertThat(document.getPages().indexOf(heading1.findDestinationPage(document))).isEqualTo(1);

            PDOutlineItem heading2 = heading1.getFirstChild();
            assertThat(heading2).isNotNull();
            assertThat(heading2.getTitle()).isEqualTo("標準化と全体最適");
        }
    }

    private static void writeWorkbook(Path path) throws Exception {
        Files.createDirectories(path.getParent());
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             OutputStream out = Files.newOutputStream(path)) {
            workbook.createSheet("First").createRow(0).createCell(0).setCellValue("first");
            workbook.setPrintArea(0, "A1:A1");
            workbook.createSheet("Second").createRow(0).createCell(0).setCellValue("second");
            workbook.setPrintArea(1, "A1:A1");
            workbook.write(out);
        }
    }
}
