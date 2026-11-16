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

package com.github.naofum.xlsx2pdf4j.internal.drawingml;

/**
 * Text rectangle insets defined by ECMA-376 preset geometry. These are
 * independent of a host document and are combined with the shape's
 * {@code bodyPr} insets by the host adapter.
 */
public final class PresetTextInsets {

    private PresetTextInsets() {
    }

    public static TextInsets forShape(PresetShape shape, double w, double h,
                                      ShapeAdjustments adj) {
        double ss = Math.min(w, h);
        ShapeAdjustments adjustments = adj == null ? ShapeAdjustments.NONE : adj;
        return switch (shape) {
            case ROUND_RECT, FLOW_CHART_ALTERNATE_PROCESS -> {
                double a = clamp(adjustments.value("adj", 16667), 0, 50000);
                double inset = ss * a / 100000d * 29289 / 100000d;
                yield new TextInsets(inset, inset, inset, inset);
            }
            case ELLIPSE, FLOW_CHART_CONNECTOR -> {
                double insetX = w * (1 - Math.sqrt(0.5)) / 2;
                double insetY = h * (1 - Math.sqrt(0.5)) / 2;
                yield new TextInsets(insetX, insetY, insetX, insetY);
            }
            case FLOW_CHART_DOCUMENT ->
                    new TextInsets(0, 0, 0, h * (1 - 17322.0 / 21600));
            case FOLDED_CORNER -> {
                double a = clamp(adjustments.value("adj", 16667), 0, 50000);
                yield new TextInsets(0, 0, 0, ss * a / 100000d);
            }
            case FLOW_CHART_MAGNETIC_DISK ->
                    new TextInsets(0, h / 3, 0, h / 6);
            case HOME_PLATE -> {
                double maxAdj = 100000d * w / ss;
                double a = clamp(adjustments.value("adj", 50000), 0, maxAdj);
                yield new TextInsets(0, 0, ss * a / 100000d / 2, 0);
            }
            case FLOW_CHART_DECISION ->
                    new TextInsets(w / 4, h / 4, w / 4, h / 4);
            case FLOW_CHART_MULTIDOCUMENT ->
                    new TextInsets(0, h * 3675 / 21600,
                            w * (1 - 18595.0 / 21600), h * (1 - 20782.0 / 21600));
            case IRREGULAR_SEAL_2 ->
                    new TextInsets(w * 5372 / 21600, h * 6382 / 21600,
                            w * (1 - 14640.0 / 21600), h * (1 - 15935.0 / 21600));
            default -> TextInsets.NONE;
        };
    }

    private static double clamp(double value, double min, double max) {
        return Math.max(min, Math.min(max, value));
    }
}
