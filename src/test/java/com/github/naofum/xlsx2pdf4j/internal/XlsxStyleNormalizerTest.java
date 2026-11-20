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

package com.github.naofum.xlsx2pdf4j.internal;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression tests for workbooks that omit applyBorder. Excel still renders
 * the referenced border, but POI reports BorderStyle.NONE until the flag is
 * normalized.
 */
class XlsxStyleNormalizerTest {

    @Test
    void appliesBorderWhenTheFlagIsOmitted() throws Exception {
        Path path = Path.of("testdata", "validation", "screen-list-ja.xlsx");
        try (InputStream in = Files.newInputStream(path);
             XSSFWorkbook workbook = new XSSFWorkbook(in)) {
            XSSFCellStyle style = workbook.getCellStyleAt(3);
            assertThat(style.getCoreXf().getBorderId()).isNotZero();
            assertThat(style.getCoreXf().isSetApplyBorder()).isFalse();
            assertThat(style.getBorderTop()).isEqualTo(BorderStyle.NONE);

            XlsxStyleNormalizer.applyExcelDefaults(workbook);

            assertThat(style.getCoreXf().getApplyBorder()).isTrue();
            assertThat(style.getBorderTop()).isEqualTo(BorderStyle.THIN);
            assertThat(style.getBorderBottom()).isEqualTo(BorderStyle.THIN);
            assertThat(style.getBorderLeft()).isEqualTo(BorderStyle.THIN);
            assertThat(style.getBorderRight()).isEqualTo(BorderStyle.THIN);
        }
    }

    @Test
    void keepsAnExplicitlyDisabledBorderFlag() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFCellStyle style = workbook.createCellStyle();
            style.setBorderTop(BorderStyle.THIN);
            style.getCoreXf().setApplyBorder(false);

            XlsxStyleNormalizer.applyExcelDefaults(workbook);

            assertThat(style.getCoreXf().isSetApplyBorder()).isTrue();
            assertThat(style.getCoreXf().getApplyBorder()).isFalse();
            assertThat(style.getBorderTop()).isEqualTo(BorderStyle.NONE);
        }
    }
}
