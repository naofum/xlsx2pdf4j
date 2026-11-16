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
import com.github.naofum.xlsx2pdf4j.api.WarningCode;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Generates the attribute-coverage workbook into target/ and verifies that it
 * still converts with the documented page count and warning set.
 */
class ValidationWorkbookTest {

    @Test
    void validationWorkbookConvertsWithExpectedPagesAndWarnings() throws Exception {
        Path workbook = Path.of("target", "validation-attributes", "attributes.xlsx");
        Path pdf = Path.of("target", "validation-attributes", "attributes.pdf");
        Files.createDirectories(pdf.getParent());

        ValidationWorkbookGenerator.generate(workbook);
        assertThat(workbook).exists().isNotEmptyFile();

        ConversionResult result = Xlsx2Pdf.convert(workbook, pdf,
                ConversionOptions.builder().dpi(150).build());

        assertThat(pdf).exists().isNotEmptyFile();
        assertThat(result.pageCount()).isEqualTo(9);
        assertThat(result.pagesBySheet())
                .containsEntry("01 Cells", 1)
                .containsEntry("02 Text", 1)
                .containsEntry("03 Shapes", 1)
                .containsEntry("04 Connectors", 1)
                .containsEntry("05 Images", 1)
                .containsEntry("06 Print", 3)
                .containsEntry("07 Warnings", 1);
        assertThat(result.warnings())
                .extracting(warning -> warning.code())
                .containsExactlyInAnyOrder(
                        WarningCode.PAGE_LIMIT,
                        WarningCode.FORMULA_NO_CACHE,
                        WarningCode.CHART_SKIPPED,
                        WarningCode.IMAGE_UNSUPPORTED,
                        WarningCode.SHAPE_APPROXIMATED);
    }
}
