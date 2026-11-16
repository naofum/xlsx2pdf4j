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

import java.awt.Shape;
import java.awt.geom.Rectangle2D;
import java.util.function.DoubleFunction;

/**
 * A format-neutral shape ready for {@link ShapePainter}. Host adapters provide
 * the frame, preset (or explicit geometry), style and text model; the painter
 * owns fill, stroke, shadow, arrows, rotation and text placement.
 */
public record DrawingShape(
        String name,
        Rectangle2D frame,
        Shape geometry,
        PresetShape preset,
        ShapeAdjustments adjustments,
        CustomGeometry customGeometry,
        DrawingStyle style,
        ShapeTextBody text,
        DoubleFunction<ShapeTextBody> textAtScale) {

    /**
     * A shape whose geometry is generated from an ECMA-376 preset.
     */
    public static DrawingShape preset(String name, Rectangle2D frame, PresetShape preset,
                                      ShapeAdjustments adjustments, DrawingStyle style,
                                      ShapeTextBody text,
                                      DoubleFunction<ShapeTextBody> textAtScale) {
        return new DrawingShape(name, frame, null, preset, adjustments, null,
                style, text, textAtScale);
    }

    /**
     * A shape whose geometry has already been built by the host adapter.
     */
    public static DrawingShape geometry(String name, Rectangle2D frame, Shape geometry,
                                        DrawingStyle style, ShapeTextBody text,
                                        DoubleFunction<ShapeTextBody> textAtScale) {
        return new DrawingShape(name, frame, geometry, null, null, null,
                style, text, textAtScale);
    }

    /**
     * A custom-geometry shape. The painter applies the per-path fill and stroke
     * flags stored in {@code customGeometry}.
     */
    public static DrawingShape custom(String name, Rectangle2D frame,
                                      CustomGeometry customGeometry, DrawingStyle style,
                                      ShapeTextBody text,
                                      DoubleFunction<ShapeTextBody> textAtScale) {
        return new DrawingShape(name, frame, null, null, null, customGeometry,
                style, text, textAtScale);
    }
}
