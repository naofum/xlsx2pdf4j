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
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Verifies that the unsupported-shape catalog covers every non-connector
 * preset that ShapeRenderer approximates, and that the generated workbook
 * converts as documented.
 */
class UnsupportedShapesWorkbookTest {

    @Test
    void catalogContainsEveryUnsupportedPresetAndConverts() throws Exception {
        List<UnsupportedShapesWorkbookGenerator.ShapeEntry> shapes =
                UnsupportedShapesWorkbookGenerator.unsupportedShapes();
        assertThat(shapes).hasSize(12);
        assertThat(shapes)
                .allSatisfy(shape -> {
                    assertThat(shape.constant()).doesNotContain("CONNECTOR");
                    assertThat(shape.preset()).isNotBlank();
                });

        Map<String, List<UnsupportedShapesWorkbookGenerator.ShapeEntry>> groups =
                UnsupportedShapesWorkbookGenerator.groupedShapes();
        assertThat(groups.get("action")).hasSize(12);

        Path workbook = Path.of("target", "unsupported-shapes", "unsupported-shapes.xlsx");
        Path pdf = Path.of("target", "unsupported-shapes", "unsupported-shapes.pdf");
        Files.createDirectories(pdf.getParent());
        UnsupportedShapesWorkbookGenerator.generate(workbook);

        ConversionResult result = Xlsx2Pdf.convert(workbook, pdf,
                ConversionOptions.builder().dpi(100).build());

        assertThat(result.pageCount()).isEqualTo(2);
        assertThat(result.pagesBySheet())
                .containsEntry("00 Index", 1)
                .containsEntry("01 Action Buttons", 1);
        assertThat(result.warnings())
                .hasSize(12)
                .allSatisfy(warning -> {
                    assertThat(warning.code()).isEqualTo(WarningCode.SHAPE_APPROXIMATED);
                    assertThat(warning.message()).contains("approximated by a rectangle");
                });
    }
}
