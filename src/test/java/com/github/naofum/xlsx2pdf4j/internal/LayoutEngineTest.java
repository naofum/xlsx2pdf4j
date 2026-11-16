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

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.nio.file.Path;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class LayoutEngineTest {

    @Test
    void explicitPrintScaleOverridesWorkbookSetting() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook(
                new FileInputStream(Path.of("testdata", "validation", "attributes.xlsx").toFile()))) {
            SheetLayout layout = SheetLayout.read(workbook, 0, workbook.getSheetAt(0), 1.0);

            List<PageLayout> pages = new LayoutEngine().layout(layout, 1, 0.5);

            assertThat(pages).isNotEmpty();
            assertThat(pages.get(0).scale).isEqualTo(0.5);
        }
    }
}
