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

import java.awt.Color;
import java.awt.Font;
import java.awt.FontMetrics;
import java.awt.Graphics2D;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;

/**
 * Renders Excel header/footer section strings with the common &amp; codes.
 */
public final class HeaderFooterRenderer {

    private static final DateTimeFormatter DATE = DateTimeFormatter.ofPattern("yyyy/M/d");
    private static final DateTimeFormatter TIME = DateTimeFormatter.ofPattern("H:mm");

    private final FontResolver fonts;

    public HeaderFooterRenderer(FontResolver fonts) {
        this.fonts = fonts;
    }

    public void render(Graphics2D g2, PageLayout page, int totalPages, String fileName,
                       java.util.List<PlacedText> textLayer) {
        drawHeaderFooter(g2, page, page.sheet.print.header, true, totalPages, fileName, textLayer);
        drawHeaderFooter(g2, page, page.sheet.print.footer, false, totalPages, fileName, textLayer);
    }

    private void drawHeaderFooter(Graphics2D g2, PageLayout page, HeaderFooterModel model,
                                  boolean header, int totalPages, String fileName,
                                  java.util.List<PlacedText> textLayer) {
        if (model == null || model.isEmpty()) {
            return;
        }
        double y = header
                ? page.sheet.print.headerMargin
                : page.pageHeight - page.sheet.print.footerMargin;
        drawSection(g2, model.left(), page, totalPages, fileName, y,
                page.sheet.print.marginLeft, Align.LEFT, header, textLayer);
        drawSection(g2, model.center(), page, totalPages, fileName, y,
                page.pageWidth / 2d, Align.CENTER, header, textLayer);
        drawSection(g2, model.right(), page, totalPages, fileName, y,
                page.pageWidth - page.sheet.print.marginRight, Align.RIGHT, header, textLayer);
    }

    private void drawSection(Graphics2D g2, String section, PageLayout page, int totalPages,
                             String fileName, double y, double anchorX, Align align, boolean header,
                             java.util.List<PlacedText> textLayer) {
        if (section == null || section.isEmpty()) {
            return;
        }
        ParsedText parsed = parse(section, page, totalPages, fileName);
        if (parsed.text.isEmpty()) {
            return;
        }
        Font font = fonts.resolveForText(parsed.fontName, parsed.fontSize, false, false, parsed.text);
        g2.setFont(font);
        g2.setColor(Color.BLACK);
        FontMetrics fm = g2.getFontMetrics(font);
        double width = fm.stringWidth(parsed.text);
        double x = switch (align) {
            case CENTER -> anchorX - width / 2d;
            case RIGHT -> anchorX - width;
            default -> anchorX;
        };
        double baseline = header ? y + fm.getAscent() : y - fm.getDescent();
        if (textLayer != null) {
            textLayer.add(new PlacedText(parsed.text, x, baseline,
                    Math.max(1, fonts.scaleSize(parsed.fontSize))));
        }
        g2.drawString(parsed.text, (float) x, (float) baseline);
    }

    private enum Align {
        LEFT, CENTER, RIGHT
    }

    private ParsedText parse(String raw, PageLayout page, int totalPages, String fileName) {
        String text = raw;
        String fontName = "MS Mincho";
        double fontSize = 10;
        StringBuilder out = new StringBuilder();
        for (int i = 0; i < text.length(); i++) {
            char ch = text.charAt(i);
            if (ch != '&') {
                out.append(ch);
                continue;
            }
            if (i + 1 >= text.length()) {
                out.append('&');
                break;
            }
            char code = text.charAt(++i);
            switch (code) {
                case '&' -> out.append('&');
                case 'L', 'C', 'R' -> {
                    // Section markers are not expected inside a section, but
                    // ignore them if present.
                }
                case 'P' -> out.append(page.pageNumber);
                case 'N' -> out.append(totalPages);
                case 'D' -> out.append(LocalDate.now().format(DATE));
                case 'T' -> out.append(LocalTime.now().format(TIME));
                case 'F' -> out.append(fileName == null ? "" : fileName);
                case 'A' -> out.append(page.sheet.name);
                case '"' -> {
                    int end = text.indexOf('"', i + 1);
                    if (end < 0) {
                        out.append('&');
                    } else {
                        String spec = text.substring(i + 1, end);
                        int comma = spec.indexOf(',');
                        fontName = comma >= 0 ? spec.substring(0, comma) : spec;
                        i = end;
                    }
                }
                default -> {
                    if (Character.isDigit(code)) {
                        int start = i;
                        while (i + 1 < text.length() && Character.isDigit(text.charAt(i + 1))) {
                            i++;
                        }
                        try {
                            fontSize = Integer.parseInt(text.substring(start, i + 1));
                        } catch (NumberFormatException ignored) {
                            // Keep the previous size.
                        }
                    } else if (code == 'B' || code == 'I' || code == 'U' || code == 'K') {
                        // Bold/italic/underline/color are ignored for now.
                    } else {
                        out.append(code);
                    }
                }
            }
        }
        return new ParsedText(out.toString(), fontName, fontSize);
    }

    private record ParsedText(String text, String fontName, double fontSize) {
    }
}
