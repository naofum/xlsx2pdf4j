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

import com.github.naofum.xlsx2pdf4j.internal.drawingml.PresetShape;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.xslf.usermodel.XSLFSimpleShape;

/**
 * Maps Apache POI presentation shape types onto the DrawingML preset names.
 */
public final class PptxPresetShapes {

    private PptxPresetShapes() {
    }

    public static PresetShape from(ShapeType type) {
        if (type == null || type == ShapeType.NOT_PRIMITIVE) {
            return null;
        }
        try {
            return PresetShape.valueOf(type.name());
        } catch (IllegalArgumentException ignored) {
            return null;
        }
    }

    public static PresetShape from(XSLFSimpleShape shape) {
        return shape == null ? null : from(shape.getShapeType());
    }
}
