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
import com.github.naofum.xlsx2pdf4j.api.PdfMetadata;
import com.github.naofum.xlsx2pdf4j.api.Pptx2Pdf;
import com.github.naofum.xlsx2pdf4j.api.PptxConversionOptions;
import com.github.naofum.xlsx2pdf4j.api.Xlsx2Pdf;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentInformation;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;

import static org.assertj.core.api.Assertions.assertThat;

class PdfMetadataTest {

    private static final Instant CREATED = Instant.parse("2026-01-02T03:04:05Z");
    private static final Instant MODIFIED = Instant.parse("2026-02-03T04:05:06Z");

    @Test
    void builderCarriesEveryMetadataField() {
        PdfMetadata metadata = PdfMetadata.builder()
                .title("Title")
                .author("Author")
                .subject("Subject")
                .keywords("alpha, beta")
                .creator("Creator")
                .producer("Producer")
                .creationDate(CREATED)
                .modificationDate(MODIFIED)
                .build();

        assertThat(metadata.title()).isEqualTo("Title");
        assertThat(metadata.author()).isEqualTo("Author");
        assertThat(metadata.subject()).isEqualTo("Subject");
        assertThat(metadata.keywords()).isEqualTo("alpha, beta");
        assertThat(metadata.creator()).isEqualTo("Creator");
        assertThat(metadata.producer()).isEqualTo("Producer");
        assertThat(metadata.creationDate()).isEqualTo(CREATED);
        assertThat(metadata.modificationDate()).isEqualTo(MODIFIED);

        PdfMetadata empty = PdfMetadata.builder().build();
        assertThat(empty.title()).isNull();
        assertThat(empty.author()).isNull();
        assertThat(empty.creationDate()).isNull();
    }

    @Test
    void defaultConversionKeepsTheExistingMetadataBehavior() throws Exception {
        Path workbook = Path.of("target", "pdf-metadata", "default.xlsx");
        Path pdf = Path.of("target", "pdf-metadata", "default.pdf");
        writeWorkbook(workbook);

        Xlsx2Pdf.convert(workbook, pdf,
                ConversionOptions.builder().dpi(100).build());

        try (PDDocument document = Loader.loadPDF(pdf.toFile())) {
            PDDocumentInformation info = document.getDocumentInformation();
            assertThat(info.getTitle()).isEqualTo("default.xlsx");
            assertThat(info.getAuthor()).isNull();
            assertThat(info.getSubject()).isNull();
            assertThat(info.getKeywords()).isNull();
            assertThat(info.getCreator()).isEqualTo("xlsx2pdf4j");
            assertThat(info.getProducer()).isEqualTo("xlsx2pdf4j");
            assertThat(info.getCreationDate()).isNull();
            assertThat(info.getModificationDate()).isNull();
        }
    }

    @Test
    void excelConversionWritesCustomMetadata() throws Exception {
        Path workbook = Path.of("target", "pdf-metadata", "excel.xlsx");
        Path pdf = Path.of("target", "pdf-metadata", "excel.pdf");
        writeWorkbook(workbook);
        PdfMetadata metadata = fullMetadata("Excel title");

        Xlsx2Pdf.convert(workbook, pdf,
                ConversionOptions.builder()
                        .dpi(100)
                        .pdfMetadata(metadata)
                        .build());

        assertMetadata(pdf, metadata);
    }

    @Test
    void powerpointConversionWritesCustomMetadata() throws Exception {
        Path pdf = Path.of("target", "pdf-metadata", "powerpoint.pdf");
        PdfMetadata metadata = fullMetadata("PowerPoint title");

        Pptx2Pdf.convert(
                Path.of("testdata", "validation", "sample1.pptx"), pdf,
                PptxConversionOptions.builder()
                        .dpi(100)
                        .pdfMetadata(metadata)
                        .build());

        assertMetadata(pdf, metadata);
    }

    @Test
    void wordConversionWritesCustomMetadata() throws Exception {
        Path pdf = Path.of("target", "pdf-metadata", "word.pdf");
        PdfMetadata metadata = fullMetadata("Word title");

        Docx2Pdf.convert(
                Path.of("testdata", "validation", "sample1.docx"), pdf,
                DocxConversionOptions.builder()
                        .dpi(100)
                        .pdfMetadata(metadata)
                        .build());

        assertMetadata(pdf, metadata);
    }

    private static PdfMetadata fullMetadata(String title) {
        return PdfMetadata.builder()
                .title(title)
                .author("Metadata Author")
                .subject("Metadata Subject")
                .keywords("metadata, test")
                .creator("Metadata Creator")
                .producer("Metadata Producer")
                .creationDate(CREATED)
                .modificationDate(MODIFIED)
                .build();
    }

    private static void assertMetadata(Path pdf, PdfMetadata metadata) throws Exception {
        try (PDDocument document = Loader.loadPDF(pdf.toFile())) {
            PDDocumentInformation info = document.getDocumentInformation();
            assertThat(info.getTitle()).isEqualTo(metadata.title());
            assertThat(info.getAuthor()).isEqualTo(metadata.author());
            assertThat(info.getSubject()).isEqualTo(metadata.subject());
            assertThat(info.getKeywords()).isEqualTo(metadata.keywords());
            assertThat(info.getCreator()).isEqualTo(metadata.creator());
            assertThat(info.getProducer()).isEqualTo(metadata.producer());
            assertThat(info.getCreationDate().toInstant()).isEqualTo(metadata.creationDate());
            assertThat(info.getModificationDate().toInstant()).isEqualTo(metadata.modificationDate());
        }
    }

    private static void writeWorkbook(Path path) throws Exception {
        Files.createDirectories(path.getParent());
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             OutputStream out = Files.newOutputStream(path)) {
            workbook.createSheet("Sheet1").createRow(0).createCell(0).setCellValue("metadata");
            workbook.write(out);
        }
    }
}
