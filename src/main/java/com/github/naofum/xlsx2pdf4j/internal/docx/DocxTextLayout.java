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

import com.github.naofum.xlsx2pdf4j.internal.FontResolver;
import com.github.naofum.xlsx2pdf4j.internal.PlacedText;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.awt.Color;
import java.awt.Font;
import java.awt.FontMetrics;
import java.awt.Graphics2D;
import java.util.ArrayList;
import java.util.List;

/**
 * Simple paragraph line breaker and painter for the Word flow layout.
 */
final class DocxTextLayout {

    record Segment(String text, Font font, Color color, boolean underline, boolean strike,
                   double width) {
    }

    record Line(List<Segment> segments, double width, double height,
                double leftIndent, double availableWidth, HorizontalAlignment alignment,
                boolean breakAfter) {
    }

    private record Token(String text, DocxStyle style, boolean pageBreak,
                         boolean whitespace, boolean cjk) {
    }

    private final FontResolver fonts = new FontResolver();

    List<Line> layout(DocxParagraphModel paragraph, double maxWidth) {
        DocxStyle style = paragraph.style();
        List<Token> tokens = tokens(paragraph);
        List<Line> lines = new ArrayList<>();
        if (tokens.isEmpty()) {
            lines.add(new Line(List.of(), 0, style.lineHeight(),
                    style.leftIndent + style.firstLineIndent, maxWidth, style.alignment, false));
            return lines;
        }

        List<Segment> segments = new ArrayList<>();
        double width = 0;
        double lineHeight = 0;
        boolean firstLine = true;
        double leftIndent = Math.max(0, style.leftIndent + (firstLine ? style.firstLineIndent : 0));
        double available = Math.max(10, maxWidth - leftIndent - Math.max(0, style.rightIndent));
        boolean breakAfter = false;

        for (Token token : tokens) {
            if (token.pageBreak()) {
                lines.add(new Line(List.copyOf(segments), width,
                        lineHeight <= 0 ? style.lineHeight() : lineHeight,
                        leftIndent, available, style.alignment, true));
                segments = new ArrayList<>();
                width = 0;
                lineHeight = 0;
                firstLine = false;
                leftIndent = Math.max(0, style.leftIndent);
                available = Math.max(10, maxWidth - leftIndent - Math.max(0, style.rightIndent));
                breakAfter = false;
                continue;
            }
            Font font = fonts.resolveForText(token.style().fontFamily,
                    token.style().fontSize, token.style().bold, token.style().italic,
                    token.text());
            double tokenWidth = font.getStringBounds(token.text(),
                    new java.awt.font.FontRenderContext(null, true, true)).getWidth();
            if (width > 0 && width + tokenWidth > available) {
                lines.add(new Line(List.copyOf(segments), width,
                        lineHeight <= 0 ? style.lineHeight() : lineHeight,
                        leftIndent, available, style.alignment, breakAfter));
                segments = new ArrayList<>();
                width = 0;
                lineHeight = 0;
                firstLine = false;
                leftIndent = Math.max(0, style.leftIndent);
                available = Math.max(10,
                        maxWidth - leftIndent - Math.max(0, style.rightIndent));
                breakAfter = false;
            }
            if (tokenWidth > available && width == 0 && !token.whitespace()) {
                String remaining = token.text();
                while (!remaining.isEmpty()) {
                    int count = 1;
                    String part = remaining.substring(0, count);
                    double partWidth = font.getStringBounds(part,
                            new java.awt.font.FontRenderContext(null, true, true)).getWidth();
                    while (count < remaining.length()) {
                        String next = remaining.substring(0, count + 1);
                        double nextWidth = font.getStringBounds(next,
                                new java.awt.font.FontRenderContext(null, true, true)).getWidth();
                        if (nextWidth > available) {
                            break;
                        }
                        count++;
                        part = next;
                        partWidth = nextWidth;
                    }
                    segments.add(new Segment(part, font, token.style().color,
                            token.style().underline, token.style().strike, partWidth));
                    width += partWidth;
                    lineHeight = Math.max(lineHeight, token.style().lineHeight());
                    remaining = remaining.substring(count);
                    if (!remaining.isEmpty()) {
                        lines.add(new Line(List.copyOf(segments), width, lineHeight,
                                leftIndent, available, style.alignment, breakAfter));
                        segments = new ArrayList<>();
                        width = 0;
                        lineHeight = 0;
                        leftIndent = Math.max(0, style.leftIndent);
                        available = Math.max(10,
                                maxWidth - leftIndent - Math.max(0, style.rightIndent));
                        breakAfter = false;
                    }
                }
                continue;
            }
            segments.add(new Segment(token.text(), font, token.style().color,
                    token.style().underline, token.style().strike, tokenWidth));
            width += tokenWidth;
            lineHeight = Math.max(lineHeight, token.style().lineHeight());
        }
        if (!segments.isEmpty() || lines.isEmpty()) {
            lines.add(new Line(List.copyOf(segments), width,
                    lineHeight <= 0 ? style.lineHeight() : lineHeight,
                    leftIndent, available, style.alignment, breakAfter));
        }
        return lines;
    }

    void draw(Graphics2D g2, Line line, double left, double top,
              List<PlacedText> textLayer) {
        double x = left;
        if (line.alignment() == HorizontalAlignment.CENTER) {
            x = left + Math.max(0, (line.availableWidth() - line.width()) / 2);
        } else if (line.alignment() == HorizontalAlignment.RIGHT) {
            x = left + Math.max(0, line.availableWidth() - line.width());
        }
        double baseline = top + line.height() * 0.82;
        double textX = x;
        double firstSize = 0;
        StringBuilder text = new StringBuilder();
        for (Segment segment : line.segments()) {
            g2.setFont(segment.font());
            g2.setColor(segment.color());
            FontMetrics metrics = g2.getFontMetrics(segment.font());
            g2.drawString(segment.text(), (float) textX, (float) baseline);
            if (segment.underline()) {
                int y = (int) Math.round(baseline + 1);
                g2.drawLine((int) Math.round(textX), y,
                        (int) Math.round(textX + segment.width()), y);
            }
            if (segment.strike()) {
                int y = (int) Math.round(baseline - metrics.getAscent() * 0.35);
                g2.drawLine((int) Math.round(textX), y,
                        (int) Math.round(textX + segment.width()), y);
            }
            if (firstSize <= 0) {
                firstSize = segment.font().getSize2D();
            }
            text.append(segment.text());
            textX += segment.width();
        }
        if (textLayer != null && !text.toString().isBlank() && firstSize > 0) {
            textLayer.add(new PlacedText(text.toString().strip(), x, baseline, firstSize));
        }
    }

    private List<Token> tokens(DocxParagraphModel paragraph) {
        List<Token> result = new ArrayList<>();
        for (DocxRunModel run : paragraph.runs()) {
            if (run.pageBreak()) {
                result.add(new Token("", run.style(), true, false, false));
            }
            if (run.text() == null || run.text().isEmpty()) {
                continue;
            }
            String text = run.text();
            int i = 0;
            while (i < text.length()) {
                char c = text.charAt(i);
                if (c == '\r') {
                    i++;
                    continue;
                }
                if (c == '\n') {
                    result.add(new Token("\n", run.style(), true, false, false));
                    i++;
                    continue;
                }
                if (c == '\t') {
                    result.add(new Token("    ", run.style(), false, true, false));
                    i++;
                    continue;
                }
                if (Character.isWhitespace(c)) {
                    result.add(new Token(" ", run.style(), false, true, false));
                    i++;
                    continue;
                }
                if (isCjk(c)) {
                    result.add(new Token(String.valueOf(c), run.style(), false, false, true));
                    i++;
                    continue;
                }
                int start = i;
                while (i < text.length()) {
                    char next = text.charAt(i);
                    if (Character.isWhitespace(next) || isCjk(next)) {
                        break;
                    }
                    i++;
                }
                result.add(new Token(text.substring(start, i), run.style(), false, false, false));
            }
        }
        return result;
    }

    private static boolean isCjk(char c) {
        Character.UnicodeBlock block = Character.UnicodeBlock.of(c);
        return block == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS
                || block == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS_EXTENSION_A
                || block == Character.UnicodeBlock.HIRAGANA
                || block == Character.UnicodeBlock.KATAKANA
                || block == Character.UnicodeBlock.HALFWIDTH_AND_FULLWIDTH_FORMS
                || block == Character.UnicodeBlock.CJK_SYMBOLS_AND_PUNCTUATION;
    }
}
