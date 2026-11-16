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

import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.Color;

/**
 * Converts POI colors (including theme and tint) to AWT colors.
 */
public final class ColorUtil {

    public static final Color BLACK = Color.BLACK;
    public static final Color WHITE = Color.WHITE;

    private ColorUtil() {
    }

    public static Color toAwt(XSSFColor color, XSSFWorkbook workbook) {
        return toAwt(color, workbook, BLACK);
    }

    public static Color toAwt(XSSFColor color, XSSFWorkbook workbook, Color fallback) {
        if (color == null) {
            return fallback;
        }
        byte[] rgb = color.getRGB();
        if (rgb == null && color.isThemed() && workbook != null && workbook.getTheme() != null) {
            XSSFColor theme = workbook.getTheme().getThemeColor(color.getTheme());
            if (theme != null) {
                rgb = theme.getRGB();
            }
        }
        // Indexed colors are normally resolved by XSSFColor#getRGB(). If not,
        // the caller's fallback is used; an explicit indexed palette can be
        // added later without changing the public API.
        if (rgb == null) {
            return fallback;
        }
        int r = rgb[0] & 0xFF;
        int g = rgb[1] & 0xFF;
        int b = rgb[2] & 0xFF;
        double tint = color.getTint();
        if (tint != 0) {
            r = applyTint(r, tint);
            g = applyTint(g, tint);
            b = applyTint(b, tint);
        }
        int alpha = rgb.length > 3 ? (rgb[3] & 0xFF) : 255;
        return new Color(r, g, b, alpha);
    }

    private static int applyTint(int channel, double tint) {
        double c = channel / 255d;
        double result = tint < 0
                ? c * (1d + tint)
                : c + (1d - c) * tint;
        return (int) Math.round(Units.clamp(result, 0d, 1d) * 255d);
    }

    public static Color withTint(Color color, double tint) {
        if (color == null || tint == 0) {
            return color;
        }
        return new Color(
                applyTint(color.getRed(), tint),
                applyTint(color.getGreen(), tint),
                applyTint(color.getBlue(), tint),
                color.getAlpha());
    }

    public static Color withAlpha(Color color, int alpha) {
        if (color == null) {
            return null;
        }
        return new Color(color.getRed(), color.getGreen(), color.getBlue(), alpha);
    }

    public static Color interpolate(Color a, Color b, double ratio) {
        double r = Units.clamp(ratio, 0d, 1d);
        return new Color(
                (int) Math.round(a.getRed() + (b.getRed() - a.getRed()) * r),
                (int) Math.round(a.getGreen() + (b.getGreen() - a.getGreen()) * r),
                (int) Math.round(a.getBlue() + (b.getBlue() - a.getBlue()) * r),
                (int) Math.round(a.getAlpha() + (b.getAlpha() - a.getAlpha()) * r));
    }
}
