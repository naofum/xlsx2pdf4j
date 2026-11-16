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
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

class JapaneseBuiltinDateFormatTest {

    @Test
    void formatsReservedJapaneseBuiltinDate() throws Exception {
        Path workbookPath = Path.of("target", "japanese-date-source.xlsx");
        Path pdfPath = Path.of("target", "japanese-date.pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Date");
            sheet.setColumnWidth(0, 3000);
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue(42566);
            CellStyle style = workbook.createCellStyle();
            // Excel built-in format 31: yyyy"年"m"月"d"日"
            style.setDataFormat((short) 31);
            cell.setCellStyle(style);
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
            assertThat(extracted).contains("2016年7月15日");
            assertThat(extracted).doesNotContain("reserved");
        }
    }
}
