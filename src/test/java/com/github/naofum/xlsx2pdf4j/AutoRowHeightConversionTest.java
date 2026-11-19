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
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * End-to-end regression for a workbook whose wrapped cell has no explicit row
 * height. The row must expand so the full text is visible in the PDF.
 */
class AutoRowHeightConversionTest {

    @Test
    void convertsWrappedCellWithoutExplicitRowHeight() throws Exception {
        Path input = Path.of("testdata", "validation", "screen-list-ja.xlsx");
        Path output = Path.of("target", "auto-row-height", "screen-list-ja.pdf");
        Files.createDirectories(output.getParent());

        ConversionResult result = Xlsx2Pdf.convert(input, output,
                ConversionOptions.builder().dpi(150).build());

        assertThat(result.pageCount()).isEqualTo(1);
        assertThat(result.warnings()).isEmpty();
        assertThat(output).exists().isNotEmptyFile();

        try (PDDocument document = Loader.loadPDF(output.toFile())) {
            String text = new PDFTextStripper().getText(document);
            assertThat(text).contains(
                    "営業課長が見積を承認する。100 万円以上は営業部長も承認する。");
        }
    }
}
