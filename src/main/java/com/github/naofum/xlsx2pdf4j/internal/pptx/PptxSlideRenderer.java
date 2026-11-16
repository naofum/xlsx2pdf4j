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

import com.github.naofum.xlsx2pdf4j.api.ConversionWarning;
import com.github.naofum.xlsx2pdf4j.api.WarningCode;
import com.github.naofum.xlsx2pdf4j.internal.PlacedText;
import com.github.naofum.xlsx2pdf4j.internal.chart.ChartRenderer;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingTheme;
import org.apache.poi.sl.draw.Drawable;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFGraphicFrame;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.util.List;

/**
 * Renders a single presentation slide with POI's Java2D renderer and overlays
 * charts that POI only draws when a fallback picture is available.
 */
public final class PptxSlideRenderer {

    private PptxSlideRenderer() {
    }

    public static BufferedImage render(XMLSlideShow presentation, XSLFSlide slide,
                                       int dpi, ChartRenderer charts,
                                       List<PlacedText> textLayer,
                                       List<ConversionWarning> warnings) {
        Dimension size = presentation.getPageSize();
        double scale = dpi / 72.0;
        int width = Math.max(1, (int) Math.round(size.getWidth() * scale));
        int height = Math.max(1, (int) Math.round(size.getHeight() * scale));
        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2 = image.createGraphics();
        try {
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING,
                    RenderingHints.VALUE_ANTIALIAS_ON);
            g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING,
                    RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
            g2.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS,
                    RenderingHints.VALUE_FRACTIONALMETRICS_ON);
            g2.setRenderingHint(RenderingHints.KEY_RENDERING,
                    RenderingHints.VALUE_RENDER_QUALITY);
            g2.setRenderingHint(RenderingHints.KEY_STROKE_CONTROL,
                    RenderingHints.VALUE_STROKE_PURE);
            g2.setRenderingHint(Drawable.FONT_HANDLER, new PptxFontManager());
            g2.setColor(Color.WHITE);
            g2.fillRect(0, 0, width, height);
            g2.scale(scale, scale);

            if (textLayer != null) {
                PptxTextLayer.collect(slide, textLayer);
            }

            try {
                slide.draw(g2);
            } catch (RuntimeException e) {
                warnings.add(new ConversionWarning(WarningCode.SHAPE_UNSUPPORTED,
                        "Slide " + slide.getSlideNumber(), null,
                        "failed to render slide: " + e.getMessage()));
            }
            drawCharts(g2, slide, charts, textLayer, warnings);
        } finally {
            g2.dispose();
        }
        return image;
    }

    private static void drawCharts(Graphics2D g2, XSLFSlide slide, ChartRenderer charts,
                                   List<PlacedText> textLayer,
                                   List<ConversionWarning> warnings) {
        DrawingTheme theme = PptxThemeReader.read(slide);
        for (XSLFShape shape : slide.getShapes()) {
            if (!(shape instanceof XSLFGraphicFrame frame)) {
                continue;
            }
            if (frame.hasChart()) {
                Rectangle2D anchor = frame.getAnchor();
                if (anchor == null || anchor.getWidth() <= 0 || anchor.getHeight() <= 0) {
                    continue;
                }
                try {
                    if (!charts.draw(g2, frame.getChart(), anchor, theme, textLayer)) {
                        warnings.add(new ConversionWarning(WarningCode.CHART_SKIPPED,
                                "Slide " + slide.getSlideNumber(), frame.getShapeName(),
                                "unsupported presentation chart"));
                    }
                } catch (RuntimeException e) {
                    warnings.add(new ConversionWarning(WarningCode.CHART_SKIPPED,
                            "Slide " + slide.getSlideNumber(), frame.getShapeName(),
                            "failed to draw chart: " + e.getMessage()));
                }
            } else if (frame.hasDiagram()) {
                warnings.add(new ConversionWarning(WarningCode.SMART_ART_SKIPPED,
                        "Slide " + slide.getSlideNumber(), frame.getShapeName(),
                        "SmartArt is not rendered"));
            }
        }
    }
}
