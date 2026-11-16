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

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.awt.Color;
import java.awt.Font;
import java.awt.FontMetrics;
import java.awt.Graphics2D;
import java.awt.Rectangle;
import java.awt.Shape;
import java.awt.font.FontRenderContext;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.text.BreakIterator;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Small text layout engine for cells and shapes. It intentionally favors
 * predictable output over a full Excel line-breaking implementation.
 */
public final class TextLayoutEngine {

    private static final double MIN_SHRINK_SIZE = 1.0;

    public record StyledRun(String text, Font font, Color color, boolean underline, boolean strike) {
    }

    public record Paragraph(List<StyledRun> runs, HorizontalAlignment horizontal, Font fallbackFont) {
    }

    /**
     * Draws shape text paragraph by paragraph. Each paragraph keeps its own
     * alignment, while anchorCtr centers the complete text block inside the
     * text rectangle. lineSpacingFactor is used by normAutofit's
     * lnSpcReduction.
     */
    public void drawParagraphs(Graphics2D g2,
                               List<Paragraph> paragraphs,
                               Rectangle2D rect,
                               VerticalAlignment vertical,
                               boolean wrap,
                               boolean anchorCtr,
                               double lineSpacingFactor,
                               int rotation) {
        if (paragraphs == null || paragraphs.isEmpty()
                || rect.getWidth() <= 0 || rect.getHeight() <= 0) {
            return;
        }
        Shape oldClip = g2.getClip();
        g2.clip(rect);
        try {
            double angle = rotationAngle(rotation);
            AffineTransform old = g2.getTransform();
            if (angle != 0) {
                g2.rotate(Math.toRadians(angle), rect.getCenterX(), rect.getCenterY());
            }
            try {
                drawParagraphsHorizontal(g2, paragraphs, rect, vertical, wrap,
                        anchorCtr, lineSpacingFactor);
            } finally {
                g2.setTransform(old);
            }
        } finally {
            g2.setClip(oldClip);
        }
    }

    private void drawParagraphsHorizontal(Graphics2D g2,
                                          List<Paragraph> paragraphs,
                                          Rectangle2D rect,
                                          VerticalAlignment vertical,
                                          boolean wrap,
                                          boolean anchorCtr,
                                          double lineSpacingFactor) {
        FontRenderContext frc = g2.getFontRenderContext();
        double spacing = Units.clamp(lineSpacingFactor, 0.1, 1.0);
        List<ParagraphLayout> layouts = new ArrayList<>();
        double widest = 0;
        double totalHeight = 0;
        for (Paragraph paragraph : paragraphs) {
            List<Token> tokens = tokenize(paragraph.runs());
            List<List<Token>> lines = layoutTokens(tokens, Math.max(1, rect.getWidth()), wrap);
            if (lines.isEmpty()) {
                lines.add(new ArrayList<>());
            }
            List<LineLayout> lineLayouts = new ArrayList<>();
            double paragraphWidth = 0;
            double paragraphHeight = 0;
            for (List<Token> line : lines) {
                double width = 0;
                double lineHeight = 0;
                Font lineFont = paragraph.fallbackFont();
                for (Token token : line) {
                    width += token.font().getStringBounds(token.text(), frc).getWidth();
                    lineHeight = Math.max(lineHeight,
                            token.font().getLineMetrics(token.text(), frc).getHeight());
                    if (lineFont == null) {
                        lineFont = token.font();
                    }
                }
                if (lineHeight <= 0 && lineFont != null) {
                    lineHeight = lineFont.getLineMetrics("Ag", frc).getHeight();
                }
                lineHeight = Math.max(1, lineHeight * spacing);
                lineLayouts.add(new LineLayout(line, width, lineHeight));
                paragraphWidth = Math.max(paragraphWidth, width);
                paragraphHeight += lineHeight;
            }
            layouts.add(new ParagraphLayout(lineLayouts, paragraph.horizontal()));
            widest = Math.max(widest, paragraphWidth);
            totalHeight += paragraphHeight;
        }

        double alignmentWidth = anchorCtr ? widest : rect.getWidth();
        double blockLeft = anchorCtr
                ? rect.getX() + (rect.getWidth() - widest) / 2d
                : rect.getX();
        double top = switch (vertical) {
            case CENTER -> rect.getY() + (rect.getHeight() - totalHeight) / 2d;
            case BOTTOM -> rect.getY() + rect.getHeight() - totalHeight;
            default -> rect.getY();
        };

        double y = top;
        for (ParagraphLayout paragraph : layouts) {
            for (LineLayout line : paragraph.lines()) {
                if (!line.tokens().isEmpty()) {
                    double x = switch (paragraph.horizontal()) {
                        case CENTER, CENTER_SELECTION ->
                                blockLeft + (alignmentWidth - line.width()) / 2d;
                        case RIGHT -> blockLeft + alignmentWidth - line.width();
                        default -> blockLeft;
                    };
                    Token first = line.tokens().get(0);
                    double baseline = y + first.font()
                            .getLineMetrics(first.text(), frc).getAscent();
                    for (Token token : line.tokens()) {
                        g2.setFont(token.font());
                        g2.setColor(token.color() == null ? Color.BLACK : token.color());
                        g2.drawString(token.text(), (float) x, (float) baseline);
                        double width = token.font().getStringBounds(token.text(), frc).getWidth();
                        if (token.underline()) {
                            g2.drawLine((int) Math.round(x), (int) Math.round(baseline + 1),
                                    (int) Math.round(x + width), (int) Math.round(baseline + 1));
                        }
                        if (token.strike()) {
                            double strikeY = baseline - token.font()
                                    .getLineMetrics(token.text(), frc).getAscent() / 2d;
                            g2.drawLine((int) Math.round(x), (int) Math.round(strikeY),
                                    (int) Math.round(x + width), (int) Math.round(strikeY));
                        }
                        x += width;
                    }
                }
                y += line.height();
            }
        }
    }

    private record LineLayout(List<Token> tokens, double width, double height) {
    }

    private record ParagraphLayout(List<LineLayout> lines,
                                   HorizontalAlignment horizontal) {
    }

    public void drawRichText(Graphics2D g2,
                             List<StyledRun> runs,
                             Rectangle2D rect,
                             HorizontalAlignment horizontal,
                             VerticalAlignment vertical,
                             boolean wrap,
                             int rotation,
                             double indentPoints) {
        if (runs == null || runs.isEmpty() || rect.getWidth() <= 0 || rect.getHeight() <= 0) {
            return;
        }
        Shape oldClip = g2.getClip();
        g2.clip(rect);
        try {
            if (rotation == 255) {
                // Rich text is rare in vertical cells; fall back to the plain
                // renderer's vertical behavior by joining the text.
                StringBuilder text = new StringBuilder();
                Font font = runs.get(0).font();
                Color color = runs.get(0).color();
                boolean underline = runs.get(0).underline();
                boolean strike = runs.get(0).strike();
                for (StyledRun run : runs) {
                    text.append(run.text());
                }
                drawVertical(g2, text.toString(), font, color, rect, horizontal, vertical, underline, strike);
                return;
            }
            double angle = rotationAngle(rotation);
            AffineTransform old = g2.getTransform();
            Rectangle2D layoutRect = rect;
            if (angle != 0) {
                layoutRect = rotatedLayoutRect(rect);
                g2.rotate(Math.toRadians(angle), rect.getCenterX(), rect.getCenterY());
            }
            try {
                drawRichHorizontal(g2, runs, layoutRect, horizontal, vertical, wrap, indentPoints);
            } finally {
                g2.setTransform(old);
            }
        } finally {
            g2.setClip(oldClip);
        }
    }

    public double richTextWidth(Graphics2D g2, List<StyledRun> runs) {
        if (runs == null || runs.isEmpty()) {
            return 0;
        }
        FontRenderContext frc = g2.getFontRenderContext();
        double widest = 0;
        double lineWidth = 0;
        for (Token token : tokenize(runs)) {
            if ("\n".equals(token.text())) {
                widest = Math.max(widest, lineWidth);
                lineWidth = 0;
            } else {
                lineWidth += token.font().getStringBounds(token.text(), frc).getWidth();
            }
        }
        return Math.max(widest, lineWidth);
    }

    public double richTextHeight(Graphics2D g2, List<StyledRun> runs,
                                 double maxWidth, boolean wrap) {
        if (runs == null || runs.isEmpty()) {
            return 0;
        }
        List<Token> tokens = tokenize(runs);
        List<List<Token>> lines = layoutTokens(tokens, Math.max(1, maxWidth), wrap);
        FontRenderContext frc = g2.getFontRenderContext();
        double total = 0;
        for (List<Token> line : lines) {
            double lineHeight = 0;
            for (Token token : line) {
                lineHeight = Math.max(lineHeight,
                        token.font().getLineMetrics(token.text(), frc).getHeight());
            }
            total += lineHeight <= 0 ? 12 : lineHeight;
        }
        return total;
    }

    private void drawRichHorizontal(Graphics2D g2,
                                    List<StyledRun> runs,
                                    Rectangle2D rect,
                                    HorizontalAlignment horizontal,
                                    VerticalAlignment vertical,
                                    boolean wrap,
                                    double indentPoints) {
        List<Token> tokens = tokenize(runs);
        double maxWidth = Math.max(1, rect.getWidth() - indentPoints);
        List<List<Token>> lines = layoutTokens(tokens, maxWidth, wrap);
        if (lines.isEmpty()) {
            return;
        }
        FontRenderContext frc = g2.getFontRenderContext();
        double totalHeight = 0;
        List<Double> heights = new ArrayList<>();
        for (List<Token> line : lines) {
            double lineHeight = 0;
            for (Token token : line) {
                lineHeight = Math.max(lineHeight, token.font().getLineMetrics(token.text(), frc).getHeight());
            }
            if (lineHeight <= 0) {
                lineHeight = 12;
            }
            heights.add(lineHeight);
            totalHeight += lineHeight;
        }
        double top = switch (vertical) {
            case CENTER -> rect.getY() + (rect.getHeight() - totalHeight) / 2d;
            case BOTTOM -> rect.getY() + rect.getHeight() - totalHeight;
            default -> rect.getY();
        };
        double y = top;
        for (int i = 0; i < lines.size(); i++) {
            List<Token> line = lines.get(i);
            if (line.isEmpty()) {
                y += heights.get(i);
                continue;
            }
            double lineWidth = 0;
            for (Token token : line) {
                lineWidth += token.font().getStringBounds(token.text(), frc).getWidth();
            }
            double x = switch (horizontal) {
                case CENTER, CENTER_SELECTION -> rect.getX() + (rect.getWidth() - lineWidth) / 2d;
                case RIGHT -> rect.getX() + rect.getWidth() - lineWidth - indentPoints;
                default -> rect.getX() + indentPoints;
            };
            double baseline = y + line.get(0).font().getLineMetrics(line.get(0).text(), frc).getAscent();
            for (Token token : line) {
                g2.setFont(token.font());
                g2.setColor(token.color() == null ? Color.BLACK : token.color());
                g2.drawString(token.text(), (float) x, (float) baseline);
                double width = token.font().getStringBounds(token.text(), frc).getWidth();
                if (token.underline()) {
                    g2.drawLine((int) Math.round(x), (int) Math.round(baseline + 1),
                            (int) Math.round(x + width), (int) Math.round(baseline + 1));
                }
                if (token.strike()) {
                    double strikeY = baseline - token.font().getLineMetrics(token.text(), frc).getAscent() / 2d;
                    g2.drawLine((int) Math.round(x), (int) Math.round(strikeY),
                            (int) Math.round(x + width), (int) Math.round(strikeY));
                }
                x += width;
            }
            y += heights.get(i);
        }
    }

    private List<Token> tokenize(List<StyledRun> runs) {
        List<Token> tokens = new ArrayList<>();
        for (StyledRun run : runs) {
            if (run.text() == null || run.text().isEmpty()) {
                continue;
            }
            String text = run.text();
            int start = 0;
            for (int i = 0; i < text.length(); i++) {
                if (text.charAt(i) == '\n') {
                    if (i > start) {
                        tokens.add(new Token(text.substring(start, i), run.font(), run.color(),
                                run.underline(), run.strike()));
                    }
                    tokens.add(new Token("\n", run.font(), run.color(), false, false));
                    start = i + 1;
                }
            }
            if (start < text.length()) {
                BreakIterator iterator = BreakIterator.getLineInstance(Locale.JAPANESE);
                iterator.setText(text);
                int a = iterator.first();
                int b = iterator.next();
                while (b != BreakIterator.DONE) {
                    String part = text.substring(a, b);
                    if (!part.isEmpty()) {
                        tokens.add(new Token(part, run.font(), run.color(), run.underline(), run.strike()));
                    }
                    a = b;
                    b = iterator.next();
                }
            }
        }
        return tokens;
    }

    private List<List<Token>> layoutTokens(List<Token> tokens, double maxWidth, boolean wrap) {
        List<List<Token>> lines = new ArrayList<>();
        List<Token> current = new ArrayList<>();
        double width = 0;
        for (Token token : tokens) {
            if ("\n".equals(token.text())) {
                lines.add(current);
                current = new ArrayList<>();
                width = 0;
                continue;
            }
            double tokenWidth = token.font().getStringBounds(token.text(),
                    new FontRenderContext(null, true, true)).getWidth();
            if (wrap && !current.isEmpty() && width + tokenWidth > maxWidth) {
                lines.add(current);
                current = new ArrayList<>();
                width = 0;
            }
            if (wrap && tokenWidth > maxWidth && current.isEmpty()) {
                for (int i = 0; i < token.text().length(); i++) {
                    String ch = token.text().substring(i, i + 1);
                    double chWidth = token.font().getStringBounds(ch,
                            new FontRenderContext(null, true, true)).getWidth();
                    if (width + chWidth > maxWidth && !current.isEmpty()) {
                        lines.add(current);
                        current = new ArrayList<>();
                        width = 0;
                    }
                    current.add(new Token(ch, token.font(), token.color(), token.underline(), token.strike()));
                    width += chWidth;
                }
                continue;
            }
            current.add(token);
            width += tokenWidth;
        }
        if (!current.isEmpty() || lines.isEmpty()) {
            lines.add(current);
        }
        return lines;
    }

    private record Token(String text, Font font, Color color, boolean underline, boolean strike) {
    }

    public double stringWidth(Graphics2D g2, String text, Font font) {
        if (text == null || text.isEmpty()) {
            return 0;
        }
        return g2.getFontMetrics(font).stringWidth(text.replace("\r", "").replace("\n", " "));
    }

    public void drawText(Graphics2D g2,
                         String text,
                         Font font,
                         Color color,
                         Rectangle2D rect,
                         HorizontalAlignment horizontal,
                         VerticalAlignment vertical,
                         boolean wrap,
                         boolean shrink,
                         int rotation,
                         double indentPoints,
                         boolean underline,
                         boolean strike) {
        if (text == null || text.isEmpty() || rect.getWidth() <= 0 || rect.getHeight() <= 0) {
            return;
        }
        String value = text.replace("\r\n", "\n").replace('\r', '\n');
        Shape oldClip = g2.getClip();
        g2.clip(rect);
        try {
            if (rotation == 255) {
                drawVertical(g2, value, font, color, rect, horizontal, vertical, underline, strike);
                return;
            }
            double angle = rotationAngle(rotation);
            if (angle != 0) {
                double cx = rect.getCenterX();
                double cy = rect.getCenterY();
                AffineTransform old = g2.getTransform();
                g2.rotate(Math.toRadians(angle), cx, cy);
                try {
                    drawHorizontal(g2, value, font, color, rotatedLayoutRect(rect), horizontal, vertical, wrap,
                            shrink, indentPoints, underline, strike);
                } finally {
                    g2.setTransform(old);
                }
            } else {
                drawHorizontal(g2, value, font, color, rect, horizontal, vertical, wrap, shrink,
                        indentPoints, underline, strike);
            }
        } finally {
            g2.setClip(oldClip);
        }
    }

    private void drawHorizontal(Graphics2D g2,
                                String text,
                                Font font,
                                Color color,
                                Rectangle2D rect,
                                HorizontalAlignment horizontal,
                                VerticalAlignment vertical,
                                boolean wrap,
                                boolean shrink,
                                double indentPoints,
                                boolean underline,
                                boolean strike) {
        Font chosen = font;
        List<String> lines = wrapLines(text, chosen, g2.getFontRenderContext(), rect.getWidth() - indentPoints, wrap);
        if (shrink) {
            double size = font.getSize2D();
            FontMetrics metrics = g2.getFontMetrics(chosen);
            while (size > MIN_SHRINK_SIZE && !fits(lines, metrics, rect, wrap)) {
                size -= 0.5;
                chosen = font.deriveFont((float) size);
                metrics = g2.getFontMetrics(chosen);
                lines = wrapLines(text, chosen, g2.getFontRenderContext(), rect.getWidth() - indentPoints, wrap);
            }
        }
        FontMetrics fm = g2.getFontMetrics(chosen);
        double lineHeight = fm.getAscent() + fm.getDescent() + fm.getLeading();
        if (lineHeight <= 0) {
            lineHeight = chosen.getSize2D();
        }
        double totalHeight = lineHeight * Math.max(1, lines.size());
        double top = switch (vertical) {
            case CENTER -> rect.getY() + (rect.getHeight() - totalHeight) / 2d;
            case BOTTOM -> rect.getY() + rect.getHeight() - totalHeight;
            default -> rect.getY();
        };
        g2.setFont(chosen);
        g2.setColor(color == null ? Color.BLACK : color);
        for (int i = 0; i < lines.size(); i++) {
            String line = lines.get(i);
            double width = fm.stringWidth(line);
            double x = switch (horizontal) {
                case CENTER, CENTER_SELECTION -> rect.getX() + (rect.getWidth() - width) / 2d;
                case RIGHT -> rect.getX() + rect.getWidth() - width - indentPoints;
                default -> rect.getX() + indentPoints;
            };
            double baseline = top + fm.getAscent() + i * lineHeight;
            g2.drawString(line, (float) x, (float) baseline);
            if (underline) {
                g2.drawLine((int) Math.round(x), (int) Math.round(baseline + 1),
                        (int) Math.round(x + width), (int) Math.round(baseline + 1));
            }
            if (strike) {
                double strikeY = baseline - fm.getAscent() / 2d;
                g2.drawLine((int) Math.round(x), (int) Math.round(strikeY),
                        (int) Math.round(x + width), (int) Math.round(strikeY));
            }
        }
    }

    /**
     * East Asian vertical layout: upright characters flow top-to-bottom and
     * columns progress right-to-left.
     */
    public void drawVerticalRichText(Graphics2D g2,
                                     List<StyledRun> runs,
                                     Rectangle2D rect,
                                     HorizontalAlignment horizontal,
                                     VerticalAlignment vertical,
                                     boolean wrap) {
        if (runs == null || runs.isEmpty() || rect.getWidth() <= 0 || rect.getHeight() <= 0) {
            return;
        }
        List<List<VChar>> columns = new ArrayList<>();
        List<VChar> current = new ArrayList<>();
        double currentHeight = 0;
        FontRenderContext frc = g2.getFontRenderContext();
        for (StyledRun run : runs) {
            if (run.text() == null) {
                continue;
            }
            for (int offset = 0; offset < run.text().length(); ) {
                int cp = run.text().codePointAt(offset);
                offset += Character.charCount(cp);
                if (cp == '\n') {
                    columns.add(current);
                    current = new ArrayList<>();
                    currentHeight = 0;
                    continue;
                }
                String ch = new String(Character.toChars(cp));
                double charHeight = run.font().getLineMetrics(ch, frc).getHeight();
                if (wrap && !current.isEmpty()
                        && currentHeight + charHeight > rect.getHeight()) {
                    columns.add(current);
                    current = new ArrayList<>();
                    currentHeight = 0;
                }
                current.add(new VChar(ch, run.font(), run.color(),
                        run.underline(), run.strike()));
                currentHeight += charHeight;
            }
        }
        columns.add(current);
        if (columns.stream().allMatch(List::isEmpty)) {
            return;
        }

        double spacing = 2;
        double totalWidth = 0;
        int nonEmpty = 0;
        for (List<VChar> column : columns) {
            double width = columnWidth(column, frc);
            totalWidth += width;
            if (!column.isEmpty()) {
                nonEmpty++;
            }
        }
        if (nonEmpty > 1) {
            totalWidth += spacing * (nonEmpty - 1);
        }
        double startX = switch (horizontal) {
            case CENTER, CENTER_SELECTION -> rect.getX() + (rect.getWidth() - totalWidth) / 2d;
            case RIGHT -> rect.getX() + rect.getWidth() - totalWidth;
            default -> rect.getX();
        };

        Shape oldClip = g2.getClip();
        g2.clip(rect);
        try {
            double columnRight = startX + totalWidth;
            for (List<VChar> column : columns) {
                double width = columnWidth(column, frc);
                if (column.isEmpty()) {
                    continue;
                }
                double height = columnHeight(column, frc);
                double y = switch (vertical) {
                    case CENTER -> rect.getY() + (rect.getHeight() - height) / 2d;
                    case BOTTOM -> rect.getY() + rect.getHeight() - height;
                    default -> rect.getY();
                };
                double x = columnRight - width;
                for (VChar ch : column) {
                    FontMetrics fm = g2.getFontMetrics(ch.font());
                    double baseline = y + fm.getAscent();
                    g2.setFont(ch.font());
                    g2.setColor(ch.color() == null ? Color.BLACK : ch.color());
                    g2.drawString(ch.text(), (float) x, (float) baseline);
                    y += ch.font().getLineMetrics(ch.text(), frc).getHeight();
                }
                columnRight = x - spacing;
            }
        } finally {
            g2.setClip(oldClip);
        }
    }

    private record VChar(String text, Font font, Color color,
                         boolean underline, boolean strike) {
    }

    private static double columnWidth(List<VChar> column, FontRenderContext frc) {
        double width = 0;
        for (VChar ch : column) {
            width = Math.max(width, ch.font().getStringBounds(ch.text(), frc).getWidth());
        }
        return width;
    }

    private static double columnHeight(List<VChar> column, FontRenderContext frc) {
        double height = 0;
        for (VChar ch : column) {
            height += ch.font().getLineMetrics(ch.text(), frc).getHeight();
        }
        return height;
    }

    private void drawVertical(Graphics2D g2,
                              String text,
                              Font font,
                              Color color,
                              Rectangle2D rect,
                              HorizontalAlignment horizontal,
                              VerticalAlignment vertical,
                              boolean underline,
                              boolean strike) {
        FontMetrics fm = g2.getFontMetrics(font);
        int count = text.codePointCount(0, text.length());
        double lineHeight = fm.getHeight();
        double totalHeight = count * lineHeight;
        double top = switch (vertical) {
            case CENTER -> rect.getY() + (rect.getHeight() - totalHeight) / 2d;
            case BOTTOM -> rect.getY() + rect.getHeight() - totalHeight;
            default -> rect.getY();
        };
        g2.setFont(font);
        g2.setColor(color == null ? Color.BLACK : color);
        int index = 0;
        for (int offset = 0; offset < text.length(); ) {
            int cp = text.codePointAt(offset);
            String ch = new String(Character.toChars(cp));
            double width = fm.stringWidth(ch);
            double x = switch (horizontal) {
                case CENTER, CENTER_SELECTION -> rect.getX() + (rect.getWidth() - width) / 2d;
                case RIGHT -> rect.getX() + rect.getWidth() - width;
                default -> rect.getX();
            };
            double baseline = top + fm.getAscent() + index * lineHeight;
            g2.drawString(ch, (float) x, (float) baseline);
            if (underline) {
                g2.drawLine((int) Math.round(x), (int) Math.round(baseline + 1),
                        (int) Math.round(x + width), (int) Math.round(baseline + 1));
            }
            if (strike) {
                double strikeY = baseline - fm.getAscent() / 2d;
                g2.drawLine((int) Math.round(x), (int) Math.round(strikeY),
                        (int) Math.round(x + width), (int) Math.round(strikeY));
            }
            index++;
            offset += Character.charCount(cp);
        }
    }

    private boolean fits(List<String> lines, FontMetrics fm, Rectangle2D rect, boolean wrap) {
        double lineHeight = fm.getAscent() + fm.getDescent() + fm.getLeading();
        if (lineHeight * lines.size() > rect.getHeight()) {
            return false;
        }
        return wrap || lines.size() <= 1;
    }

    public List<String> wrapLines(String text, Font font, FontRenderContext frc, double maxWidth, boolean wrap) {
        List<String> result = new ArrayList<>();
        if (text == null) {
            return result;
        }
        for (String paragraph : text.split("\n", -1)) {
            if (!wrap) {
                result.add(paragraph);
                continue;
            }
            if (paragraph.isEmpty()) {
                result.add("");
                continue;
            }
            BreakIterator iterator = BreakIterator.getLineInstance(Locale.JAPANESE);
            iterator.setText(paragraph);
            int start = iterator.first();
            int lastFit = start;
            int boundary = iterator.next();
            while (boundary != BreakIterator.DONE) {
                String candidate = paragraph.substring(start, boundary);
                if (font.getStringBounds(candidate, frc).getWidth() > maxWidth && start < boundary) {
                    if (lastFit > start) {
                        result.add(paragraph.substring(start, lastFit));
                        start = lastFit;
                    } else {
                        result.add(candidate);
                        start = boundary;
                    }
                    iterator.setText(paragraph);
                    lastFit = start;
                    boundary = iterator.following(start);
                } else {
                    lastFit = boundary;
                    boundary = iterator.next();
                }
            }
            if (start < paragraph.length()) {
                result.add(paragraph.substring(start));
            } else if (result.isEmpty()) {
                result.add("");
            }
        }
        return result;
    }

    private double rotationAngle(int rotation) {
        if (rotation <= 0 || rotation == 255) {
            return 0;
        }
        return Math.min(180, rotation);
    }

    /**
     * Rotated text is laid out in a rectangle with its width and height
     * swapped so that it can use the full cell height as its text direction.
     * The rectangle keeps the original center and is clipped to the cell, so
     * centered rotated text stays inside the cell instead of starting at the
     * unrotated left edge and being clipped away.
     */
    private Rectangle2D rotatedLayoutRect(Rectangle2D rect) {
        double cx = rect.getCenterX();
        double cy = rect.getCenterY();
        return new Rectangle2D.Double(cx - rect.getHeight() / 2d,
                cy - rect.getWidth() / 2d,
                rect.getHeight(),
                rect.getWidth());
    }
}
