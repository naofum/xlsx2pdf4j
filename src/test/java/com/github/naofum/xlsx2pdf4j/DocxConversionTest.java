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

import com.github.naofum.xlsx2pdf4j.api.Docx2Pdf;
import com.github.naofum.xlsx2pdf4j.api.Docx2PdfException;
import com.github.naofum.xlsx2pdf4j.api.DocxConversionOptions;
import com.github.naofum.xlsx2pdf4j.api.DocxConversionResult;
import com.github.naofum.xlsx2pdf4j.api.ErrorCode;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class DocxConversionTest {

    @Test
    void convertsSampleDocuments() throws Exception {
        Path outputDir = Path.of("target", "docx");
        Files.createDirectories(outputDir);
        String[] names = {"sample1", "sample2"};
        int[] pages = {3, 2};
        for (int i = 0; i < names.length; i++) {
            DocxConversionResult result = Docx2Pdf.convert(
                    Path.of("testdata", "validation", names[i] + ".docx"),
                    outputDir.resolve(names[i] + ".pdf"),
                    DocxConversionOptions.builder().dpi(150).build());
            assertThat(result.pageCount()).isEqualTo(pages[i]);
            assertThat(result.warnings()).isEmpty();
            assertThat(result.output()).exists();
            assertThat(result.outputSize()).isLessThan(50L * 1024 * 1024);
        }
    }

    @Test
    void writesSearchableTextLayer() throws Exception {
        Path outputDir = Path.of("target", "docx");
        Files.createDirectories(outputDir);
        String[] names = {"sample1", "sample2"};
        String[][] expected = {
                {"New Platform", "サービスのコンセプト"},
                {"機能比較", "顧客管理", "取引先管理"}
        };
        for (int i = 0; i < names.length; i++) {
            Path pdf = outputDir.resolve(names[i] + "-text.pdf");
            Docx2Pdf.convert(Path.of("testdata", "validation", names[i] + ".docx"), pdf,
                    DocxConversionOptions.builder().dpi(100).build());
            try (PDDocument document = Loader.loadPDF(pdf.toFile())) {
                String text = new PDFTextStripper().getText(document);
                for (String value : expected[i]) {
                    assertThat(text).contains(value);
                }
            }
        }
    }

    @Test
    void rejectsUnsupportedInputFormat() {
        assertThatThrownBy(() -> Docx2Pdf.convert(
                Path.of("README.md"),
                Path.of("target", "docx", "invalid.pdf")))
                .isInstanceOf(Docx2PdfException.class)
                .extracting(e -> ((Docx2PdfException) e).code())
                .isEqualTo(ErrorCode.INPUT_FORMAT);
    }

    @Test
    void enforcesPageLimit() {
        assertThatThrownBy(() -> Docx2Pdf.convert(
                Path.of("testdata", "validation", "sample1.docx"),
                Path.of("target", "docx", "limit.pdf"),
                DocxConversionOptions.builder().maxPages(1).build()))
                .isInstanceOf(Docx2PdfException.class)
                .extracting(e -> ((Docx2PdfException) e).code())
                .isEqualTo(ErrorCode.LIMIT_EXCEEDED);
    }
}
