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

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import com.github.naofum.xlsx2pdf4j.internal.TextLayoutEngine;

import java.util.List;
import java.util.Objects;

/**
 * Format-neutral shape text model. Host adapters extract paragraphs and runs
 * from their own text body, then supply the shape insets and body properties.
 */
public record ShapeTextBody(
        List<TextLayoutEngine.Paragraph> paragraphs,
        List<TextLayoutEngine.StyledRun> runs,
        String text,
        TextInsets insets,
        VerticalAlignment vertical,
        boolean wrap,
        boolean anchorCtr,
        boolean eastAsianVertical,
        boolean normAutofit,
        boolean fontScaleSet,
        double lineSpacing) {

    public boolean hasDistinctParagraphAlignment() {
        if (paragraphs.isEmpty()) {
            return false;
        }
        HorizontalAlignment first = paragraphs.get(0).horizontal();
        for (int i = 1; i < paragraphs.size(); i++) {
            if (paragraphs.get(i).horizontal() != first) {
                return true;
            }
        }
        return false;
    }

    public boolean hasNonCenteredParagraph() {
        for (TextLayoutEngine.Paragraph paragraph : paragraphs) {
            if (paragraph.horizontal() != HorizontalAlignment.CENTER
                    && paragraph.horizontal() != HorizontalAlignment.CENTER_SELECTION) {
                return true;
            }
        }
        return false;
    }

    public boolean uniformStyle() {
        if (runs.isEmpty()) {
            return false;
        }
        TextLayoutEngine.StyledRun first = runs.get(0);
        for (int i = 1; i < runs.size(); i++) {
            TextLayoutEngine.StyledRun run = runs.get(i);
            if (!first.font().equals(run.font())
                    || !Objects.equals(first.color(), run.color())
                    || first.underline() != run.underline()
                    || first.strike() != run.strike()) {
                return false;
            }
        }
        return true;
    }
}
