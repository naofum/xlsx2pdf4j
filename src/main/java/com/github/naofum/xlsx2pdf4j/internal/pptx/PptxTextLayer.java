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

import com.github.naofum.xlsx2pdf4j.internal.PlacedText;
import org.apache.poi.sl.usermodel.Insets2D;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import java.awt.geom.Rectangle2D;
import java.util.List;

/**
 * Builds an approximate invisible text layer from slide text shapes. The
 * positions are deliberately simple: searchability is the goal, not exact
 * line-by-line placement.
 */
public final class PptxTextLayer {

    private static final double DEFAULT_FONT_SIZE = 18.0;

    private PptxTextLayer() {
    }

    public static void collect(XSLFSlide slide, List<PlacedText> target) {
        if (slide == null || target == null) {
            return;
        }
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextShape text) {
                collect(text, target);
            }
        }
    }

    private static void collect(XSLFTextShape text, List<PlacedText> target) {
        Rectangle2D anchor = text.getAnchor();
        if (anchor == null || anchor.getWidth() <= 0 || anchor.getHeight() <= 0) {
            return;
        }
        Insets2D insets = text.getInsets();
        double left = anchor.getX() + (insets == null ? 0 : insets.left);
        double top = anchor.getY() + (insets == null ? 0 : insets.top);
        double baseline = top;
        for (XSLFTextParagraph paragraph : text.getTextParagraphs()) {
            double size = paragraph.getDefaultFontSize() == null
                    ? DEFAULT_FONT_SIZE : paragraph.getDefaultFontSize();
            for (XSLFTextRun run : paragraph.getTextRuns()) {
                if (run.getFontSize() != null && run.getFontSize() > 0) {
                    size = run.getFontSize();
                    break;
                }
            }
            String value = paragraph.getText();
            if (value == null || value.isEmpty()) {
                baseline += size * 1.2;
                continue;
            }
            for (String line : value.split("\\R", -1)) {
                if (!line.isEmpty()) {
                    target.add(new PlacedText(line, left, baseline + size * 0.8, size));
                }
                baseline += size * 1.2;
            }
        }
    }
}
