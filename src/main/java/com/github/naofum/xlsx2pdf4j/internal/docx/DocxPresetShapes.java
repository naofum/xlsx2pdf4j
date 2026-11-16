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

import com.github.naofum.xlsx2pdf4j.internal.drawingml.PresetShape;

/**
 * Maps DrawingML {@code a:prstGeom/@prst} names (camelCase) onto the
 * format-neutral {@link PresetShape} enum.
 */
public final class DocxPresetShapes {

    private DocxPresetShapes() {
    }

    public static PresetShape fromOoxml(String name) {
        if (name == null || name.isBlank()) {
            return null;
        }
        try {
            return PresetShape.valueOf(enumName(name));
        } catch (IllegalArgumentException ignored) {
            return null;
        }
    }

    static String enumName(String name) {
        StringBuilder out = new StringBuilder(name.length() + 8);
        char previous = 0;
        for (int i = 0; i < name.length(); i++) {
            char c = name.charAt(i);
            if (i > 0 && Character.isUpperCase(c)
                    && (Character.isLowerCase(previous) || Character.isDigit(previous))) {
                out.append('_');
            } else if (i > 0 && Character.isDigit(c) && Character.isLetter(previous)) {
                out.append('_');
            }
            out.append(Character.toUpperCase(c));
            previous = c;
        }
        return out.toString();
    }
}
