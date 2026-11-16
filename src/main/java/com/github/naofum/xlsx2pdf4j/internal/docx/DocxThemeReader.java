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

package com.github.naofum.xlsx2pdf4j.internal.docx;

import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingTheme;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTheme;

import java.io.InputStream;

/**
 * Reads a WordprocessingML theme part into the shared DrawingML theme.
 */
public final class DocxThemeReader {

    private DocxThemeReader() {
    }

    public static DrawingTheme read(XWPFDocument document) {
        if (document == null) {
            return DrawingTheme.defaults();
        }
        XWPFTheme theme = document.getTheme();
        if (theme == null || theme.getPackagePart() == null) {
            return DrawingTheme.defaults();
        }
        try (InputStream input = theme.getPackagePart().getInputStream()) {
            return DrawingTheme.fromThemeXml(input);
        } catch (Exception ignored) {
            return DrawingTheme.defaults();
        }
    }
}
