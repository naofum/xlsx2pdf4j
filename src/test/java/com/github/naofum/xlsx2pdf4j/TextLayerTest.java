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
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

class TextLayerTest {

    @Test
    void extractsFullTextEvenWhenTheRasterCellClipsIt() throws Exception {
        String fullText = "クリップされても検索できる長い日本語テキスト"
                + "これはセル幅を超えるため画面では途中で切れるがテキスト層には全文を残す";
        Path workbookPath = Path.of("target", "text-layer-source.xlsx");
        Path pdfPath = Path.of("target", "text-layer.pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Text");
            sheet.setColumnWidth(0, 900);
            Row row = sheet.createRow(0);
            row.setHeightInPoints(15);
            Cell cell = row.createCell(0);
            cell.setCellValue(fullText);
            workbook.setPrintArea(0, "A1:A1");
            try (OutputStream out = Files.newOutputStream(workbookPath)) {
                workbook.write(out);
            }
        }

        Xlsx2Pdf.convert(workbookPath, pdfPath,
                ConversionOptions.builder()
                        .textLayerMode(TextLayerMode.INVISIBLE)
                        .build());

        try (PDDocument document = Loader.loadPDF(pdfPath.toFile())) {
            String extracted = new PDFTextStripper().getText(document);
            assertThat(extracted).contains(fullText);
        }

        Path noTextPath = Path.of("target", "text-layer-none.pdf");
        Xlsx2Pdf.convert(workbookPath, noTextPath,
                ConversionOptions.builder()
                        .textLayerMode(TextLayerMode.NONE)
                        .build());
        try (PDDocument document = Loader.loadPDF(noTextPath.toFile())) {
            String extracted = new PDFTextStripper().getText(document);
            assertThat(extracted).doesNotContain(fullText);
        }
    }
}
