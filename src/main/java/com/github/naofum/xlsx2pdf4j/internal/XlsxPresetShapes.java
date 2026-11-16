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

import com.github.naofum.xlsx2pdf4j.internal.drawingml.PresetShape;
import org.apache.poi.ss.usermodel.ShapeTypes;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

/**
 * Maps Apache POI spreadsheet shape ids onto the format-neutral DrawingML
 * preset names. PowerPoint and Word adapters can map their own shape enums to
 * {@link PresetShape} in the same way.
 */
final class XlsxPresetShapes {

    private static final Map<Integer, PresetShape> BY_TYPE = new HashMap<>();

    static {
        for (Field field : ShapeTypes.class.getFields()) {
            if (field.getType() != int.class) {
                continue;
            }
            try {
                PresetShape preset = PresetShape.valueOf(field.getName());
                BY_TYPE.put(field.getInt(null), preset);
            } catch (IllegalArgumentException | IllegalAccessException ignored) {
                // ShapeTypes contains a few constants without a preset name.
            }
        }
    }

    private XlsxPresetShapes() {
    }

    static PresetShape from(int shapeType) {
        return BY_TYPE.get(shapeType);
    }
}
