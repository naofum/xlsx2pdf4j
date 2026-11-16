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

package com.github.naofum.xlsx2pdf4j.internal.pptx;

import com.github.naofum.xlsx2pdf4j.internal.BundledFonts;
import org.apache.poi.common.usermodel.fonts.FontInfo;
import org.apache.poi.sl.draw.DrawFontManagerDefault;

import java.awt.Font;
import java.awt.Graphics2D;
import java.util.Locale;

/**
 * Maps presentation fonts to the bundled Noto fonts. Theme fonts in PPTX are
 * often resolved to logical AWT fonts that lack Japanese glyphs on some
 * systems, so a deterministic fallback is required.
 */
final class PptxFontManager extends DrawFontManagerDefault {

    static final String SANS = "Noto Sans JP";
    static final String SERIF = "Noto Serif JP";

    @Override
    public FontInfo getMappedFont(Graphics2D graphics, FontInfo fontInfo) {
        return new PptxFontInfo(family(fontInfo));
    }

    @Override
    public FontInfo getFallbackFont(Graphics2D graphics, FontInfo fontInfo) {
        return new PptxFontInfo(family(fontInfo));
    }

    @Override
    public Font createAWTFont(Graphics2D graphics, FontInfo fontInfo,
                              double size, boolean bold, boolean italic) {
        int style = (bold ? Font.BOLD : 0) | (italic ? Font.ITALIC : 0);
        Font base = SERIF.equals(family(fontInfo)) ? BundledFonts.serif() : BundledFonts.sans();
        if (base == null) {
            base = new Font(Font.SANS_SERIF, style, 12);
        }
        return base.deriveFont(style, (float) size);
    }

    private static String family(FontInfo fontInfo) {
        if (fontInfo == null || fontInfo.getTypeface() == null) {
            return SANS;
        }
        String name = fontInfo.getTypeface().toLowerCase(Locale.ROOT);
        if (name.contains("明朝") || name.contains("mincho") || name.contains("serif")) {
            return SERIF;
        }
        return SANS;
    }
}
