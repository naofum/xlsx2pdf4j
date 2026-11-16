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
import java.awt.FontFormatException;
import java.awt.GraphicsEnvironment;
import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Loads the permissively licensed Japanese fonts shipped in the jar. They are
 * fallbacks for machines that do not have the Windows Japanese fonts.
 */
public final class BundledFonts {

    private static final Map<String, Font> FONTS = loadFonts();
    private static final byte[] SANS_BYTES = readResource("/fonts/NotoSansJP.ttf");
    private static final byte[] SERIF_BYTES = readResource("/fonts/NotoSerifJP.ttf");

    private BundledFonts() {
    }

    public static byte[] sansBytes() {
        return SANS_BYTES == null ? null : SANS_BYTES.clone();
    }

    public static byte[] serifBytes() {
        return SERIF_BYTES == null ? null : SERIF_BYTES.clone();
    }

    public static Font sans() {
        return FONTS.get("Noto Sans JP");
    }

    public static Font serif() {
        return FONTS.get("Noto Serif JP");
    }

    public static boolean isAvailable() {
        return !FONTS.isEmpty();
    }

    private static Map<String, Font> loadFonts() {
        Map<String, Font> fonts = new LinkedHashMap<>();
        load("/fonts/NotoSansJP.ttf", fonts);
        load("/fonts/NotoSerifJP.ttf", fonts);
        return fonts;
    }

    private static byte[] readResource(String resource) {
        try (InputStream in = BundledFonts.class.getResourceAsStream(resource)) {
            return in == null ? null : in.readAllBytes();
        } catch (IOException | RuntimeException e) {
            return null;
        }
    }

    private static void load(String resource, Map<String, Font> fonts) {
        try (InputStream in = BundledFonts.class.getResourceAsStream(resource)) {
            if (in == null) {
                return;
            }
            Font font = Font.createFont(Font.TRUETYPE_FONT, in);
            GraphicsEnvironment.getLocalGraphicsEnvironment().registerFont(font);
            fonts.put(font.getFamily(), font);
        } catch (FontFormatException | IOException | RuntimeException ignored) {
            // Missing or unsupported bundled font: system fonts remain usable.
        }
    }
}
