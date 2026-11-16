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
 * Format-neutral text insets in points.
 */
public record TextInsets(double left, double top, double right, double bottom) {

    public static final TextInsets NONE = new TextInsets(0, 0, 0, 0);

    public TextInsets plus(TextInsets other) {
        return new TextInsets(left + other.left, top + other.top,
                right + other.right, bottom + other.bottom);
    }
}
