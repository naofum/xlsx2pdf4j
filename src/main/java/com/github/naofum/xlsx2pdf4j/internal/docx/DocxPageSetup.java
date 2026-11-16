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

import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.math.BigDecimal;

/**
 * Page geometry in PDF points.
 */
public record DocxPageSetup(double width, double height,
                            double topMargin, double bottomMargin,
                            double leftMargin, double rightMargin,
                            double headerDistance, double footerDistance) {

    private static final double A4_WIDTH = 595.35;
    private static final double A4_HEIGHT = 842.0;
    private static final double DEFAULT_MARGIN = 72.0;

    public static DocxPageSetup from(CTSectPr section) {
        double width = A4_WIDTH;
        double height = A4_HEIGHT;
        if (section != null && section.getPgSz() != null) {
            CTPageSz size = section.getPgSz();
            width = points(size.getW(), A4_WIDTH);
            height = points(size.getH(), A4_HEIGHT);
            if (size.getOrient() != null && size.getOrient().toString().contains("landscape")
                    && width < height) {
                double swap = width;
                width = height;
                height = swap;
            }
        }
        double top = DEFAULT_MARGIN;
        double bottom = DEFAULT_MARGIN;
        double left = DEFAULT_MARGIN;
        double right = DEFAULT_MARGIN;
        double header = 36;
        double footer = 18;
        if (section != null && section.getPgMar() != null) {
            CTPageMar margin = section.getPgMar();
            top = points(margin.getTop(), top);
            bottom = points(margin.getBottom(), bottom);
            left = points(margin.getLeft(), left);
            right = points(margin.getRight(), right);
            header = points(margin.getHeader(), header);
            footer = points(margin.getFooter(), footer);
        }
        return new DocxPageSetup(width, height, top, bottom, left, right, header, footer);
    }

    public double contentWidth() {
        return width - leftMargin - rightMargin;
    }

    private static double points(Object twips, double fallback) {
        if (twips == null) {
            return fallback;
        }
        try { return new BigDecimal(twips.toString()).doubleValue() / 20.0; } catch (NumberFormatException ignored) { return fallback; }
    }
}
