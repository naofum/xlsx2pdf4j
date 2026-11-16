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

import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingTheme;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.awt.Color;
import java.io.FileInputStream;
import java.nio.file.Path;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class ThemeColorsTest {

    @Test
    void missingThemeFallsBackToDefaults() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook(
                new FileInputStream(Path.of("testdata", "validation", "attributes.xlsx").toFile()))) {
            DrawingTheme theme = ThemeColors.read(workbook);
            Map<String, Color> colors = theme.colors();

            assertThat(colors).isEqualTo(DrawingTheme.defaults().colors());
            assertThat(colors.get("accent1")).isEqualTo(new Color(0x4F, 0x81, 0xBD));
            assertThat(colors.get("bg1")).isEqualTo(Color.WHITE);
        }
    }
}
