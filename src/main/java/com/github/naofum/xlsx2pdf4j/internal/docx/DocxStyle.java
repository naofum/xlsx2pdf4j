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

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPrBase;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTUnderline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLineSpacingRule;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;

import java.awt.Color;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.List;

/**
 * Resolved paragraph and run formatting for WordprocessingML.
 */
final class DocxStyle {

    private static final String DEFAULT_FONT = "MS Mincho";
    private static final double DEFAULT_SIZE = 11.0;

    String fontFamily = DEFAULT_FONT;
    double fontSize = DEFAULT_SIZE;
    boolean bold;
    boolean italic;
    boolean underline;
    boolean strike;
    Color color = Color.BLACK;
    HorizontalAlignment alignment = HorizontalAlignment.LEFT;
    double spaceBefore;
    double spaceAfter;
    double lineHeight;
    double lineMultiplier = 1.0;
    double leftIndent;
    double rightIndent;
    double firstLineIndent;
    boolean pageBreakBefore;
    Integer numberingId;
    Integer numberingLevel;

    DocxStyle copy() {
        DocxStyle copy = new DocxStyle();
        copy.fontFamily = fontFamily;
        copy.fontSize = fontSize;
        copy.bold = bold;
        copy.italic = italic;
        copy.underline = underline;
        copy.strike = strike;
        copy.color = color;
        copy.alignment = alignment;
        copy.spaceBefore = spaceBefore;
        copy.spaceAfter = spaceAfter;
        copy.lineHeight = lineHeight;
        copy.lineMultiplier = lineMultiplier;
        copy.leftIndent = leftIndent;
        copy.rightIndent = rightIndent;
        copy.firstLineIndent = firstLineIndent;
        copy.pageBreakBefore = pageBreakBefore;
        copy.numberingId = numberingId;
        copy.numberingLevel = numberingLevel;
        return copy;
    }

    void applyRun(CTRPr properties) {
        if (properties == null) {
            return;
        }
        if (properties.sizeOfRFontsArray() > 0) {
            CTFonts fonts = properties.getRFontsArray(0);
            String family = fonts.getAscii();
            if (family == null || family.isBlank()) {
                family = fonts.getEastAsia();
            }
            if (family == null || family.isBlank()) {
                family = fonts.getHAnsi();
            }
            if (family != null && !family.isBlank()) {
                fontFamily = family;
            }
        }
        if (properties.sizeOfSzArray() > 0) {
            Object value = properties.getSzArray(0).getVal();
            double size = number(value) / 2.0;
            if (size > 0) {
                fontSize = size;
            }
        }
        if (!properties.getBList().isEmpty()) {
            bold = on(properties.getBArray(0));
        }
        if (!properties.getIList().isEmpty()) {
            italic = on(properties.getIArray(0));
        }
        if (!properties.getUList().isEmpty()) {
            underline = underline(properties.getUArray(0));
        }
        if (!properties.getStrikeList().isEmpty()) {
            strike = on(properties.getStrikeArray(0));
        }
        if (!properties.getColorList().isEmpty()) {
            Color value = color(properties.getColorArray(0));
            if (value != null) {
                color = value;
            }
        }
    }

    void applyParagraph(CTPPrBase properties) {
        if (properties == null) {
            return;
        }
        if (properties.getJc() != null) {
            String value = String.valueOf(properties.getJc().getVal());
            if (value.contains("center")) {
                alignment = HorizontalAlignment.CENTER;
            } else if (value.contains("right")) {
                alignment = HorizontalAlignment.RIGHT;
            } else if (value.contains("both") || value.contains("justify")
                    || value.contains("distribute")) {
                alignment = HorizontalAlignment.JUSTIFY;
            } else {
                alignment = HorizontalAlignment.LEFT;
            }
        }
        CTSpacing spacing = properties.getSpacing();
        if (spacing != null) {
            if (spacing.getBefore() != null) {
                spaceBefore = number(spacing.getBefore()) / 20.0;
            }
            if (spacing.getAfter() != null) {
                spaceAfter = number(spacing.getAfter()) / 20.0;
            }
            if (spacing.getLine() != null) {
                double line = number(spacing.getLine());
                STLineSpacingRule.Enum rule = spacing.getLineRule();
                if (rule == STLineSpacingRule.AUTO || rule == null) {
                    if (line > 0) {
                        lineMultiplier = line / 240.0;
                    }
                } else {
                    lineHeight = line / 20.0;
                }
            }
        }
        CTInd indentation = properties.getInd();
        if (indentation != null) {
            if (indentation.getLeft() != null) {
                leftIndent = number(indentation.getLeft()) / 20.0;
            }
            if (indentation.getRight() != null) {
                rightIndent = number(indentation.getRight()) / 20.0;
            }
            if (indentation.getFirstLine() != null) {
                firstLineIndent = number(indentation.getFirstLine()) / 20.0;
            }
            if (indentation.getHanging() != null) {
                firstLineIndent = -number(indentation.getHanging()) / 20.0;
            }
        }
        if (properties.getNumPr() != null) {
            if (properties.getNumPr().getNumId() != null
                    && properties.getNumPr().getNumId().getVal() != null) {
                numberingId = properties.getNumPr().getNumId().getVal().intValue();
            }
            if (properties.getNumPr().getIlvl() != null
                    && properties.getNumPr().getIlvl().getVal() != null) {
                numberingLevel = properties.getNumPr().getIlvl().getVal().intValue();
            }
        }
        if (properties.getPageBreakBefore() != null) {
            pageBreakBefore = on(properties.getPageBreakBefore());
        }
    }

    double lineHeight() {
        if (lineHeight > 0) {
            return lineHeight;
        }
        double base = Math.max(fontSize, 1) * 1.6;
        return base * Math.max(0.5, lineMultiplier);
    }

    private static boolean on(CTOnOff value) {
        if (value == null || value.getVal() == null) {
            return true;
        }
        String text = value.getVal().toString();
        return !"0".equals(text) && !"false".equalsIgnoreCase(text);
    }

    private static boolean underline(CTUnderline value) {
        if (value == null || value.getVal() == null) {
            return false;
        }
        return value.getVal() != STUnderline.NONE;
    }

    private static Color color(CTColor value) {
        if (value == null || value.getVal() == null) {
            return null;
        }
        Object raw = value.getVal();
        if (raw instanceof byte[] bytes && bytes.length >= 3) {
            return new Color(bytes[0] & 0xFF, bytes[1] & 0xFF, bytes[2] & 0xFF);
        }
        String text = raw.toString();
        if (text.length() != 6 || "auto".equalsIgnoreCase(text)) {
            return null;
        }
        try {
            return new Color(Integer.parseInt(text, 16));
        } catch (NumberFormatException ignored) {
            return null;
        }
    }

    private static double number(Object value) {
        if (value == null) {
            return 0;
        }
        if (value instanceof Number number) {
            return number.doubleValue();
        }
        if (value instanceof BigInteger integer) {
            return integer.doubleValue();
        }
        try {
            return new BigDecimal(value.toString()).doubleValue();
        } catch (NumberFormatException ignored) {
            return 0;
        }
    }

    private static <T> T first(List<T> list) {
        return list == null || list.isEmpty() ? null : list.get(0);
    }
}
