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

import com.github.naofum.xlsx2pdf4j.internal.BundledFonts;
import org.apache.poi.common.usermodel.fonts.FontGroup;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSheet;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 * Replaces theme and text-run fonts with the bundled Noto fonts before
 * rendering. POI does not always resolve the effective theme font for a text
 * run, which produces missing-glyph boxes for Japanese titles.
 */
public final class PptxFontNormalizer {

    private PptxFontNormalizer() {
    }

    public static void normalize(XMLSlideShow presentation) {
        // Trigger BundledFonts static initialization and font registration.
        BundledFonts.sans();
        for (XSLFSlide slide : presentation.getSlides()) {
            normalizeSheet(slide);
            if (slide.getSlideLayout() != null) {
                normalizeSheet(slide.getSlideLayout());
            }
            if (slide.getSlideMaster() != null) {
                normalizeSheet(slide.getSlideMaster());
            }
        }
    }

    private static void normalizeSheet(XSLFSheet sheet) {
        for (XSLFShape shape : sheet.getShapes()) {
            normalizeShape(shape);
        }
    }

    private static void normalizeShape(XSLFShape shape) {
        if (shape instanceof XSLFTextShape text) {
            normalizeText(text);
        } else if (shape instanceof XSLFTable table) {
            for (XSLFTableRow row : table.getRows()) {
                for (XSLFTableCell cell : row.getCells()) {
                    normalizeText(cell);
                }
            }
        } else if (shape instanceof XSLFGroupShape group) {
            for (XSLFShape child : group.getShapes()) {
                normalizeShape(child);
            }
        }
    }

    private static void normalizeText(XSLFTextShape text) {
        for (XSLFTextParagraph paragraph : text.getTextParagraphs()) {
            for (XSLFTextRun run : paragraph.getTextRuns()) {
                run.setFontFamily(PptxFontManager.SANS, FontGroup.LATIN);
                run.setFontFamily(PptxFontManager.SANS, FontGroup.EAST_ASIAN);
            }
        }
    }
}
