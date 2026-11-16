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

import java.awt.Font;
import java.text.Normalizer;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Maps Excel font names to AWT fonts. The first version uses Windows fonts and
 * falls back to logical fonts; bundled Noto fonts can be added to the resource
 * directory without changing callers.
 */
public final class FontResolver {

    private static final Pattern ESCAPED = Pattern.compile("_x([0-9A-Fa-f]{4})_");
    private static final String DEFAULT_SANS = Font.SANS_SERIF;
    private static final String DEFAULT_SERIF = Font.SERIF;

    private final Map<String, Font> cache = new HashMap<>();
    private final Map<String, String> familyCache = new HashMap<>();
    private final double scale;

    public FontResolver() {
        this(1.0);
    }

    public FontResolver(double scale) {
        this.scale = scale > 0 ? scale : 1.0;
    }

    public double scaleSize(double sizePoints) {
        return sizePoints * scale;
    }

    public Font resolve(String excelName, double sizePoints, boolean bold, boolean italic) {
        String family = family(excelName);
        int style = (bold ? Font.BOLD : Font.PLAIN) | (italic ? Font.ITALIC : Font.PLAIN);
        double scaledSize = scaleSize(sizePoints);
        String key = family + '\u0000' + style + '\u0000' + scaledSize;
        return cache.computeIfAbsent(key, ignored -> new Font(family, style, 1).deriveFont((float) scaledSize));
    }

    public Font resolve(String excelName, double sizePoints, boolean bold, boolean italic, boolean underline, boolean strike) {
        return resolve(excelName, sizePoints, bold, italic);
    }

    /**
     * Resolve a font that can display all characters in the supplied text.
     * Shape text often inherits Calibri/Arial while containing Japanese; those
     * fonts do not provide Japanese glyphs and Java2D would paint '?'.
     */
    public Font resolveForText(String excelName, double sizePoints,
                               boolean bold, boolean italic, String text) {
        Font font = resolve(excelName, sizePoints, bold, italic);
        if (text == null || text.isEmpty() || font.canDisplayUpTo(text) < 0) {
            return font;
        }
        String preferred = isSerif(excelName) ? "MS PMincho" : "MS PGothic";
        Font japanese = resolve(preferred, sizePoints, bold, italic);
        if (japanese.canDisplayUpTo(text) < 0) {
            return japanese;
        }
        Font bundled = resolve(isSerif(excelName) ? "Noto Serif JP" : "Noto Sans JP",
                sizePoints, bold, italic);
        if (bundled.canDisplayUpTo(text) < 0) {
            return bundled;
        }
        // Last resort: let Java2D choose a logical font with wider coverage.
        return new Font(Font.SANS_SERIF, style(bold, italic), 1)
                .deriveFont((float) scaleSize(sizePoints));
    }

    private static int style(boolean bold, boolean italic) {
        return (bold ? Font.BOLD : Font.PLAIN) | (italic ? Font.ITALIC : Font.PLAIN);
    }

    private static boolean isSerif(String excelName) {
        String name = excelName == null ? "" : decodeEscapes(excelName).toLowerCase(Locale.ROOT);
        return name.contains("mincho") || name.contains("明朝") || name.contains("serif");
    }

    public String family(String excelName) {
        String key = excelName == null ? "" : excelName;
        return familyCache.computeIfAbsent(key, this::mapFamily);
    }

    private String mapFamily(String raw) {
        if (raw == null || raw.isBlank()) {
            return DEFAULT_SANS;
        }
        String decoded = decodeEscapes(raw);
        String normalized = Normalizer.normalize(decoded, Normalizer.Form.NFKC).trim();
        String lower = normalized.toLowerCase(Locale.ROOT);

        if (lower.contains("wingdings")) {
            // Wingdings glyphs are handled by WingdingsMapper; use a normal
            // sans font for the replacement characters.
            return DEFAULT_SANS;
        }
        if (lower.contains("mincho") || normalized.contains("明朝")) {
            return select("MS PMincho", "MS Mincho", bundledSerif(), DEFAULT_SERIF);
        }
        if (lower.contains("gothic") || lower.contains("ゴシック") || normalized.contains("中ゴシック")) {
            return select("MS PGothic", "MS Gothic", bundledSans(), DEFAULT_SANS);
        }
        if (lower.contains("meiryo") || normalized.contains("メイリオ")) {
            return "Meiryo";
        }
        if (lower.contains("yu gothic") || normalized.contains("游ゴシック")) {
            return "Yu Gothic";
        }
        if (lower.contains("yu mincho") || normalized.contains("游明朝")) {
            return "Yu Mincho";
        }
        if (lower.contains("carlito")) {
            return "Carlito";
        }
        if (lower.contains("arial")) {
            return "Arial";
        }
        if (lower.contains("times new roman")) {
            return "Times New Roman";
        }
        if (lower.contains("courier")) {
            return "Courier New";
        }
        if (lower.contains("verdana")) {
            return "Verdana";
        }
        if (lower.contains("century schoolbook")) {
            return "Century Schoolbook";
        }
        if (lower.contains("ms sans serif") || lower.contains("helv") || lower.contains("tms rmn")) {
            return DEFAULT_SANS;
        }
        if (lower.contains("hg") || normalized.contains("丸ゴシック")) {
            return "MS PGothic";
        }
        return normalized;
    }

    private static String select(String first, String second, String bundled, String fallback) {
        if (isAvailable(first)) {
            return first;
        }
        if (isAvailable(second)) {
            return second;
        }
        if (bundled != null && isAvailable(bundled)) {
            return bundled;
        }
        return fallback;
    }

    private static boolean isAvailable(String family) {
        if (family == null || family.isBlank()) {
            return false;
        }
        Font font = new Font(family, Font.PLAIN, 12);
        String actual = font.getFamily();
        // Unknown logical names resolve to Dialog. A requested Japanese family
        // may resolve to its localized family name, which is still a success.
        return !Font.DIALOG.equalsIgnoreCase(actual)
                || family.equalsIgnoreCase(font.getFontName())
                || actual.toLowerCase(Locale.ROOT).contains(family.toLowerCase(Locale.ROOT));
    }

    private static String bundledSans() {
        return BundledFonts.isAvailable() && BundledFonts.sans() != null
                ? BundledFonts.sans().getFamily()
                : null;
    }

    private static String bundledSerif() {
        return BundledFonts.isAvailable() && BundledFonts.serif() != null
                ? BundledFonts.serif().getFamily()
                : null;
    }

    static String decodeEscapes(String value) {
        Matcher matcher = ESCAPED.matcher(value);
        StringBuilder out = new StringBuilder(value.length());
        int last = 0;
        while (matcher.find()) {
            out.append(value, last, matcher.start());
            try {
                int codePoint = Integer.parseInt(matcher.group(1), 16);
                out.appendCodePoint(codePoint);
            } catch (NumberFormatException e) {
                out.append(matcher.group());
            }
            last = matcher.end();
        }
        out.append(value, last, value.length());
        return out.toString();
    }
}
