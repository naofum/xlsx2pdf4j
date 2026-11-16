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

import com.github.naofum.xlsx2pdf4j.internal.FontResolver;
import com.github.naofum.xlsx2pdf4j.internal.TextLayoutEngine;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingShape;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingStyle;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingTheme;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.PresetShape;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.PresetTextInsets;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.ShapeAdjustments;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.ShapeTextBody;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.TextInsets;
import org.apache.poi.sl.draw.geom.GuideIf;
import org.apache.poi.sl.usermodel.ColorStyle;
import org.apache.poi.sl.usermodel.Insets2D;
import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.xddf.usermodel.text.XDDFAutoFit;
import org.apache.poi.xddf.usermodel.text.XDDFTextBody;
import org.apache.poi.xslf.usermodel.XSLFSimpleShape;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import java.awt.Color;
import java.awt.Font;
import java.awt.geom.Rectangle2D;
import java.util.ArrayList;
import java.util.List;
import java.util.function.DoubleFunction;

/**
 * Adapts POI presentation shapes to the format-neutral DrawingML core.
 */
public final class PptxShapeAdapter {

    private PptxShapeAdapter() {
    }

    public static PresetShape presetShape(XSLFSimpleShape shape) {
        return PptxPresetShapes.from(shape);
    }

    public static ShapeAdjustments adjustments(XSLFSimpleShape shape) {
        return (name, fallback) -> {
            if (shape == null) {
                return fallback;
            }
            GuideIf guide = shape.getAdjustValue(name);
            if (guide == null || guide.getFmla() == null) {
                return fallback;
            }
            String formula = guide.getFmla().trim();
            if (!formula.startsWith("val ")) {
                return fallback;
            }
            try {
                return Double.parseDouble(formula.substring(4).trim());
            } catch (NumberFormatException ignored) {
                return fallback;
            }
        };
    }

    public static DrawingStyle style(XSLFSimpleShape shape, DrawingTheme theme) {
        if (shape == null || shape.getXmlObject() == null) {
            return new DrawingStyle();
        }
        DrawingTheme colors = theme == null ? DrawingTheme.defaults() : theme;
        return DrawingStyle.parse(shape.getXmlObject().xmlText(), colors.colors());
    }

    public static Rectangle2D frame(XSLFSimpleShape shape, double scale) {
        if (shape == null) {
            return null;
        }
        Rectangle2D anchor = shape.getAnchor();
        double factor = scale <= 0 ? 1.0 : scale;
        return new Rectangle2D.Double(anchor.getX() * factor, anchor.getY() * factor,
                anchor.getWidth() * factor, anchor.getHeight() * factor);
    }

    /**
     * Builds the neutral model consumed by {@link com.github.naofum.xlsx2pdf4j.internal.drawingml.ShapePainter}.
     */
    public static DrawingShape drawingShape(XSLFSimpleShape shape, DrawingTheme theme,
                                            FontResolver fonts, double layoutScale) {
        if (shape == null) {
            return null;
        }
        double scale = layoutScale <= 0 ? 1.0 : layoutScale;
        XSLFTextShape textShape = shape instanceof XSLFTextShape text ? text : null;
        ShapeTextBody text = textShape == null
                ? null : textBody(textShape, theme, fonts, scale);
        DoubleFunction<ShapeTextBody> textAtScale = textShape == null ? null
                : additionalScale -> textBody(
                        textShape, theme, fonts, scale, additionalScale);
        return DrawingShape.preset(shape.getShapeName(), frame(shape, scale),
                presetShape(shape), adjustments(shape), style(shape, theme), text,
                textAtScale);
    }

    public static ShapeTextBody textBody(XSLFTextShape shape, DrawingTheme theme,
                                         FontResolver fonts, double layoutScale) {
        return textBody(shape, theme, fonts, layoutScale, 1.0);
    }

    public static ShapeTextBody textBody(XSLFTextShape shape, DrawingTheme theme,
                                         FontResolver fonts, double layoutScale,
                                         double additionalScale) {
        if (shape == null) {
            return null;
        }
        List<XSLFTextParagraph> paragraphs = shape.getTextParagraphs();
        if (paragraphs.isEmpty()) {
            return null;
        }
        double scale = layoutScale <= 0 ? 1.0 : layoutScale;
        double extra = additionalScale <= 0 ? 1.0 : additionalScale;
        double fontSizeScale = scale * extra;
        Rectangle2D anchor = shape.getAnchor();
        PresetShape preset = presetShape(shape);
        ShapeAdjustments adjust = adjustments(shape);
        TextInsets presetInsets = preset == null ? TextInsets.NONE
                : PresetTextInsets.forShape(preset, anchor.getWidth(), anchor.getHeight(), adjust);
        TextInsets bodyInsets = bodyInsets(shape.getInsets());
        TextInsets insets = presetInsets.plus(bodyInsets);
        DrawingTheme colors = theme == null ? DrawingTheme.defaults() : theme;

        Font baseFont = fonts.resolve("Arial", 18 * fontSizeScale, false, false);
        List<TextLayoutEngine.Paragraph> layoutParagraphs = new ArrayList<>();
        List<TextLayoutEngine.StyledRun> runs = new ArrayList<>();
        StringBuilder allText = new StringBuilder();
        boolean hasText = false;

        for (int i = 0; i < paragraphs.size(); i++) {
            XSLFTextParagraph paragraph = paragraphs.get(i);
            Double defaultSize = paragraph.getDefaultFontSize();
            String defaultFamily = paragraph.getDefaultFontFamily();
            List<TextLayoutEngine.StyledRun> paragraphRuns = new ArrayList<>();
            for (XSLFTextRun run : paragraph.getTextRuns()) {
                String text = run.getRawText();
                if (text == null || text.isEmpty()) {
                    continue;
                }
                double size = run.getFontSize() == null || run.getFontSize() <= 0
                        ? (defaultSize == null || defaultSize <= 0 ? 18 : defaultSize)
                        : run.getFontSize();
                String family = run.getFontFamily();
                if (family == null || family.isBlank()) {
                    family = defaultFamily;
                }
                if (family == null || family.isBlank()) {
                    family = "Arial";
                }
                boolean bold = run.isBold();
                boolean italic = run.isItalic();
                boolean underline = run.isUnderlined();
                boolean strike = run.isStrikethrough();
                Color color = fontColor(run, colors);
                Font font = fonts.resolveForText(family, size * fontSizeScale, bold, italic, text);
                paragraphRuns.add(new TextLayoutEngine.StyledRun(
                        text, font, color, underline, strike));
                allText.append(text);
                hasText = true;
            }
            Font fallback = paragraphRuns.isEmpty() ? baseFont : paragraphRuns.get(0).font();
            layoutParagraphs.add(new TextLayoutEngine.Paragraph(
                    paragraphRuns, horizontal(paragraph.getTextAlign()), fallback));
            runs.addAll(paragraphRuns);
            if (i + 1 < paragraphs.size()) {
                Font newlineFont = runs.isEmpty() ? baseFont : runs.get(0).font();
                Color newlineColor = runs.isEmpty()
                        ? colors.color("tx1") : runs.get(0).color();
                runs.add(new TextLayoutEngine.StyledRun("\n", newlineFont,
                        newlineColor == null ? Color.BLACK : newlineColor, false, false));
                allText.append('\n');
            }
        }
        if (!hasText || runs.isEmpty()) {
            return null;
        }

        XDDFTextBody textBody = shape.getTextBody();
        XDDFAutoFit autoFit = textBody == null || textBody.getBodyProperties() == null
                ? null : textBody.getBodyProperties().getAutoFit();
        boolean normAutofit = shape.getTextAutofit() == TextShape.TextAutofit.NORMAL;
        boolean fontScaleSet = autoFit != null && autoFit.getFontScale() > 0;
        double lineSpacing = 1.0;
        if (autoFit != null && autoFit.getLineSpaceReduction() > 0) {
            lineSpacing = 1.0 - autoFit.getLineSpaceReduction() / 100000d;
        }
        TextShape.TextDirection direction = shape.getTextDirection();
        boolean eastAsianVertical = direction == TextShape.TextDirection.VERTICAL
                || direction == TextShape.TextDirection.STACKED;
        return new ShapeTextBody(layoutParagraphs, runs, allText.toString(), insets,
                vertical(shape.getVerticalAlignment()), shape.getWordWrap(),
                shape.isHorizontalCentered(), eastAsianVertical,
                normAutofit, fontScaleSet, lineSpacing);
    }

    private static TextInsets bodyInsets(Insets2D insets) {
        if (insets == null) {
            return TextInsets.NONE;
        }
        return new TextInsets(insets.left, insets.top, insets.right, insets.bottom);
    }

    private static Color fontColor(XSLFTextRun run, DrawingTheme theme) {
        PaintStyle paint = run.getFontColor();
        if (paint instanceof PaintStyle.SolidPaint solid) {
            ColorStyle style = solid.getSolidColor();
            if (style != null && style.getColor() != null) {
                return style.getColor();
            }
        }
        Color color = theme.color("tx1");
        return color == null ? Color.BLACK : color;
    }

    private static HorizontalAlignment horizontal(TextParagraph.TextAlign align) {
        if (align == null) {
            return HorizontalAlignment.LEFT;
        }
        return switch (align) {
            case CENTER -> HorizontalAlignment.CENTER;
            case RIGHT -> HorizontalAlignment.RIGHT;
            case JUSTIFY, JUSTIFY_LOW, DIST, THAI_DIST -> HorizontalAlignment.JUSTIFY;
            default -> HorizontalAlignment.LEFT;
        };
    }

    private static VerticalAlignment vertical(
            org.apache.poi.sl.usermodel.VerticalAlignment vertical) {
        if (vertical == null) {
            return VerticalAlignment.TOP;
        }
        return switch (vertical) {
            case MIDDLE, JUSTIFIED, DISTRIBUTED -> VerticalAlignment.CENTER;
            case BOTTOM -> VerticalAlignment.BOTTOM;
            default -> VerticalAlignment.TOP;
        };
    }
}
