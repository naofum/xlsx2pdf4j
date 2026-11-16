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

/**
 * Unit conversions. All layout calculations use points (1/72 inch).
 */
public final class Units {

    public static final double EMU_PER_INCH = 914_400d;
    public static final double EMU_PER_POINT = 12_700d;
    public static final double POINTS_PER_INCH = 72d;
    public static final double EXCEL_PIXELS_PER_INCH = 96d;

    private Units() {
    }

    public static double emuToPoints(long emu) {
        return emu / EMU_PER_POINT;
    }

    public static double inchesToPoints(double inches) {
        return inches * POINTS_PER_INCH;
    }

    public static double pixelsToPoints(double pixels) {
        return pixels * POINTS_PER_INCH / EXCEL_PIXELS_PER_INCH;
    }

    public static int pointsToPixels(double points, int dpi) {
        return Math.max(1, (int) Math.round(points * dpi / POINTS_PER_INCH));
    }

    /**
     * Paper dimensions in points, portrait orientation. Unknown sizes fall back to A4.
     */
    public static double[] paperSizePoints(int paperSizeCode) {
        return switch (paperSizeCode) {
            case 1 -> new double[]{612, 792};    // Letter
            case 3 -> new double[]{792, 1224};   // Tabloid
            case 5 -> new double[]{612, 1008};   // Legal
            case 8 -> new double[]{842, 1191};   // A3
            case 9 -> new double[]{595, 842};    // A4
            case 11 -> new double[]{420, 595};   // A5
            case 12 -> new double[]{728, 1032};  // B4
            case 13 -> new double[]{516, 729};   // B5
            default -> new double[]{595, 842};   // A4
        };
    }

    public static double clamp(double value, double min, double max) {
        return Math.max(min, Math.min(max, value));
    }
}
