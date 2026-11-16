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

import com.github.naofum.xlsx2pdf4j.internal.PlacedText;
import com.github.naofum.xlsx2pdf4j.internal.TextLayoutEngine;

import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.util.ArrayList;
import java.util.List;

/**
 * Paints a {@link ShapeTextBody} through the shared text layout engine. Host
 * adapters only have to build the model; paragraph wrapping, alignment,
 * vertical text and normAutofit scaling stay in one place.
 */
public final class ShapeTextRenderer {

    private ShapeTextRenderer() {
    }

    public static void draw(Graphics2D g2, Rectangle2D frame, ShapeTextBody body,
                            TextLayoutEngine engine, List<PlacedText> textLayer) {
        if (body == null || body.runs().isEmpty()) {
            return;
        }
        Rectangle2D textRect = textRect(frame, body.insets());
        if (textLayer != null) {
            TextLayoutEngine.StyledRun first = body.runs().get(0);
            double fontSize = Math.max(1, first.font().getSize2D());
            textLayer.add(new PlacedText(body.text(), textRect.getX(),
                    textRect.getY() + Math.min(textRect.getHeight(), fontSize + 1),
                    fontSize));
        }

        if (body.eastAsianVertical()) {
            engine.drawVerticalRichText(g2, body.runs(), textRect,
                    body.paragraphs().get(0).horizontal(), body.vertical(), body.wrap());
            return;
        }

        boolean anchorShift = body.anchorCtr() && body.hasNonCenteredParagraph();
        boolean paragraphLayout = anchorShift
                || body.normAutofit()
                || body.hasDistinctParagraphAlignment();
        if (paragraphLayout) {
            engine.drawParagraphs(g2, body.paragraphs(), textRect, body.vertical(),
                    body.wrap(), anchorShift, body.lineSpacing(), 0);
        } else if (body.uniformStyle()) {
            TextLayoutEngine.StyledRun first = body.runs().get(0);
            engine.drawText(g2, body.text(), first.font(), first.color(), textRect,
                    body.paragraphs().get(0).horizontal(), body.vertical(),
                    body.wrap(), false, 0, 0,
                    first.underline(), first.strike());
        } else {
            engine.drawRichText(g2, body.runs(), textRect,
                    body.paragraphs().get(0).horizontal(), body.vertical(),
                    body.wrap(), 0, 0);
        }
    }

    public static Rectangle2D textRect(Rectangle2D frame, TextInsets insets) {
        return new Rectangle2D.Double(
                frame.getX() + insets.left(),
                frame.getY() + insets.top(),
                Math.max(1, frame.getWidth() - insets.left() - insets.right()),
                Math.max(1, frame.getHeight() - insets.top() - insets.bottom()));
    }

    /**
     * Returns the largest scale not exceeding 1.0 that makes the text fit the
     * shape. The adapter rebuilds the text model at that scale so font
     * fallback is resolved for the final size.
     */
    public static double fitScale(Graphics2D g2, Rectangle2D frame,
                                  ShapeTextBody body, TextLayoutEngine engine) {
        if (body == null || body.runs().isEmpty()) {
            return 1.0;
        }
        Rectangle2D textRect = textRect(frame, body.insets());
        if (fits(g2, textRect, body, body.runs(), engine, 1.0)) {
            return 1.0;
        }
        double low = 0.1;
        double high = 1.0;
        for (int i = 0; i < 12; i++) {
            double middle = (low + high) / 2d;
            if (fits(g2, textRect, body, body.runs(), engine, middle)) {
                low = middle;
            } else {
                high = middle;
            }
        }
        return low;
    }

    private static boolean fits(Graphics2D g2, Rectangle2D textRect,
                                ShapeTextBody body,
                                List<TextLayoutEngine.StyledRun> runs,
                                TextLayoutEngine engine, double scale) {
        List<TextLayoutEngine.StyledRun> scaled = scale == 1.0
                ? runs : scaleRuns(runs, scale);
        if (!body.wrap()
                && engine.richTextWidth(g2, scaled) > textRect.getWidth() + 0.5) {
            return false;
        }
        double height = engine.richTextHeight(g2, scaled, textRect.getWidth(), body.wrap())
                * body.lineSpacing();
        return height <= textRect.getHeight() + 0.5;
    }

    private static List<TextLayoutEngine.StyledRun> scaleRuns(
            List<TextLayoutEngine.StyledRun> runs, double scale) {
        List<TextLayoutEngine.StyledRun> scaled = new ArrayList<>();
        for (TextLayoutEngine.StyledRun run : runs) {
            Font font = run.font().deriveFont((float) (run.font().getSize2D() * scale));
            scaled.add(new TextLayoutEngine.StyledRun(run.text(), font, run.color(),
                    run.underline(), run.strike()));
        }
        return scaled;
    }
}
