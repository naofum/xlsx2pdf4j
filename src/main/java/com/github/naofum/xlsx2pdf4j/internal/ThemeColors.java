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

import java.io.InputStream;

/**
 * Excel adapter for the format-neutral DrawingML theme. Other hosts can read
 * their own theme part and call {@link DrawingTheme#fromThemeXml(InputStream)}.
 */
final class ThemeColors {

    private ThemeColors() {
    }

    static DrawingTheme read(XSSFWorkbook workbook) {
        if (workbook == null || workbook.getTheme() == null) {
            return DrawingTheme.defaults();
        }
        try (InputStream input = workbook.getTheme().getPackagePart().getInputStream()) {
            return DrawingTheme.fromThemeXml(input);
        } catch (Exception ignored) {
            return DrawingTheme.defaults();
        }
    }
}
