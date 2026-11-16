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
import com.github.naofum.xlsx2pdf4j.internal.Units;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.LinearGradientPaint;
import java.awt.Paint;
import java.awt.RadialGradientPaint;
import java.awt.Shape;
import java.awt.Stroke;
import java.awt.geom.AffineTransform;
import java.awt.geom.Area;
import java.awt.geom.Ellipse2D;
import java.awt.geom.Path2D;
import java.awt.geom.Point2D;
import java.awt.geom.Rectangle2D;
import java.util.ArrayList;
import java.util.List;

/**
 * The one Java2D rendering path for DrawingML shapes. Excel, PPTX and DOCX
 * adapters only build a {@link DrawingShape}; fill, stroke, shadow, arrows,
 * rotation/flip and text layout stay here.
 */
public final class ShapePainter {

    private ShapePainter() {
    }

    /**
     * An additional filled or stroked path that is not part of the main
     * geometry. Excel uses this for folded corners, callout leaders, bevel
     * faces and similar preset details.
     */
    public record Fill(Shape shape, String mode) {
    }

    /**
     * Host-specific decorations around the main geometry. Presentation and
     * word-processing shapes normally use {@link #NONE}.
     */
    public record Extras(
            Shape shadowGeometry,
            Shape outline,
            Shape extraFill,
            List<Fill> fills,
            List<Shape> detailOutlines,
            boolean strokeOnly,
            boolean lineMarker,
            boolean lineShape) {

        public static final Extras NONE = new Extras(null, null, null,
                List.of(), List.of(), false, false, false);

        public Extras {
            fills = fills == null ? List.of() : List.copyOf(fills);
            detailOutlines = detailOutlines == null ? List.of() : List.copyOf(detailOutlines);
        }
    }

    public static boolean draw(Graphics2D g2, DrawingShape shape,
                               TextLayoutEngine textEngine, List<PlacedText> textLayer) {
        return draw(g2, shape, Extras.NONE, textEngine, textLayer);
    }

    public static boolean draw(Graphics2D g2, DrawingShape shape, Extras extras,
                               TextLayoutEngine textEngine, List<PlacedText> textLayer) {
        if (g2 == null || shape == null || shape.frame() == null) {
            return false;
        }
        Rectangle2D frame = shape.frame();
        boolean lineShape = extras.lineShape()
                || shape.preset() == PresetShape.LINE
                || shape.preset() == PresetShape.STRAIGHT_CONNECTOR_1;
        if ((frame.getWidth() <= 0 && frame.getHeight() <= 0)
                || (!lineShape && (frame.getWidth() <= 0 || frame.getHeight() <= 0))) {
            return false;
        }
        DrawingStyle style = shape.style() == null ? new DrawingStyle() : shape.style();
        CustomGeometry custom = shape.customGeometry();
        Shape geometry = shape.geometry();
        if (geometry == null && shape.preset() != null) {
            geometry = PresetGeometry.create(shape.preset(),
                    frame.getX(), frame.getY(), frame.getWidth(), frame.getHeight(),
                    shape.adjustments() == null
                            ? ShapeAdjustments.NONE : shape.adjustments());
        }
        if (custom == null && geometry == null) {
            return false;
        }

        AffineTransform oldTransform = g2.getTransform();
        double cx = frame.getCenterX();
        double cy = frame.getCenterY();
        if (style.rotationDegrees != 0 || style.flipH || style.flipV) {
            g2.rotate(Math.toRadians(style.rotationDegrees), cx, cy);
            if (style.flipH) {
                g2.translate(cx, 0);
                g2.scale(-1, 1);
                g2.translate(-cx, 0);
            }
            if (style.flipV) {
                g2.translate(0, cy);
                g2.scale(1, -1);
                g2.translate(0, -cy);
            }
        }
        try {
            if (custom != null) {
                drawCustomGeometry(g2, custom, style);
            } else {
                paintGeometry(g2, shape, extras, style, geometry);
            }
            drawText(g2, shape, textEngine, textLayer);
        } finally {
            g2.setTransform(oldTransform);
        }
        return true;
    }

    private static void paintGeometry(Graphics2D g2, DrawingShape shape, Extras extras,
                                      DrawingStyle style, Shape geometry) {
        if (style.shadow != null) {
            Shape shadowGeometry = extras.shadowGeometry() == null
                    ? geometry : extras.shadowGeometry();
            drawShadow(g2, shadowGeometry, style.shadow);
        }
        if (!style.fillNone && !extras.strokeOnly()) {
            g2.setPaint(fillPaint(style, geometry.getBounds2D()));
            g2.fill(geometry);
            if (extras.extraFill() != null) {
                g2.setColor(darkenLess(fillColor(style)));
                g2.fill(extras.extraFill());
            }
            for (Fill fill : extras.fills()) {
                g2.setColor(fillModeColor(fill.mode(), fillColor(style)));
                g2.fill(fill.shape());
            }
        }
        if (!style.lineNone) {
            g2.setColor(style.line == null ? Color.BLACK : style.line);
            if (!extras.lineMarker()) {
                drawOutline(g2, extras.outline() == null ? geometry : extras.outline(), style);
            }
            for (Shape detail : extras.detailOutlines()) {
                drawOutline(g2, detail, style);
            }
            boolean lineShape = extras.lineShape()
                    || shape.preset() == PresetShape.LINE
                    || shape.preset() == PresetShape.STRAIGHT_CONNECTOR_1;
            if (lineShape) {
                Rectangle2D frame = shape.frame();
                double x1 = frame.getX();
                double y1 = frame.getY();
                double x2 = x1 + frame.getWidth();
                double y2 = y1 + frame.getHeight();
                if (style.headEnd != null) {
                    drawArrowHead(g2, x1, y1, x2, y2, style.headEnd, style.lineWidth);
                }
                if (style.tailEnd != null) {
                    drawArrowHead(g2, x2, y2, x1, y1, style.tailEnd, style.lineWidth);
                }
            }
        }
    }

    private static void drawText(Graphics2D g2, DrawingShape shape,
                                 TextLayoutEngine textEngine, List<PlacedText> textLayer) {
        ShapeTextBody body = shape.text();
        if (body == null || body.runs().isEmpty() || textEngine == null) {
            return;
        }
        Rectangle2D frame = shape.frame();
        if (body.normAutofit() && !body.fontScaleSet() && shape.textAtScale() != null) {
            double scale = ShapeTextRenderer.fitScale(g2, frame, body, textEngine);
            if (Math.abs(scale - 1.0) > 0.0001) {
                body = shape.textAtScale().apply(scale);
                if (body == null || body.runs().isEmpty()) {
                    return;
                }
            }
        }
        ShapeTextRenderer.draw(g2, frame, body, textEngine, textLayer);
    }

    private static void drawCustomGeometry(Graphics2D g2, CustomGeometry custom,
                                           DrawingStyle style) {
        if (style.shadow != null) {
            Area shadow = new Area();
            for (CustomPath path : custom.paths()) {
                if (path.fill()) {
                    shadow.add(new Area(path.path()));
                }
            }
            if (!shadow.isEmpty()) {
                drawShadow(g2, shadow, style.shadow);
            }
        }
        if (!style.fillNone) {
            for (CustomPath path : custom.paths()) {
                if (!path.fill()) {
                    continue;
                }
                String mode = path.fillMode();
                if (mode == null || mode.isBlank() || "norm".equals(mode)) {
                    g2.setPaint(fillPaint(style, path.path().getBounds2D()));
                } else if ("darkenLess".equals(mode)) {
                    g2.setColor(darkenLess(fillColor(style)));
                } else if ("darken".equals(mode)) {
                    g2.setColor(darken(fillColor(style), 0.6));
                } else if ("lightenLess".equals(mode)) {
                    g2.setColor(lighten(fillColor(style), 0.2));
                } else if ("lighten".equals(mode)) {
                    g2.setColor(lighten(fillColor(style), 0.4));
                } else {
                    g2.setPaint(fillPaint(style, path.path().getBounds2D()));
                }
                g2.fill(path.path());
            }
        }
        if (!style.lineNone) {
            g2.setColor(style.line == null ? Color.BLACK : style.line);
            for (CustomPath path : custom.paths()) {
                if (path.stroke()) {
                    drawOutline(g2, path.path(), style);
                }
            }
        }
    }

    private static void drawShadow(Graphics2D g2, Shape geometry, DrawingStyle.Shadow shadow) {
        Shape shifted = AffineTransform.getTranslateInstance(shadow.dx(), shadow.dy())
                .createTransformedShape(geometry);
        g2.setColor(shadow.color());
        if (shadow.blur() > 0.1) {
            Stroke old = g2.getStroke();
            g2.setStroke(new BasicStroke((float) shadow.blur(),
                    BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
            g2.draw(shifted);
            g2.setStroke(old);
        }
        g2.fill(shifted);
    }

    private static Paint fillPaint(DrawingStyle style, Rectangle2D bounds) {
        List<DrawingStyle.GradientStop> stops = style.gradientStops;
        if (stops == null || stops.size() < 2) {
            return style.fill == null ? Color.WHITE : style.fill;
        }
        List<DrawingStyle.GradientStop> normalized = normalizeStops(stops);
        if (normalized.size() < 2) {
            return normalized.isEmpty() ? Color.WHITE : normalized.get(0).color();
        }
        float[] fractions = new float[normalized.size()];
        Color[] colors = new Color[normalized.size()];
        for (int i = 0; i < normalized.size(); i++) {
            fractions[i] = (float) normalized.get(i).position();
            colors[i] = normalized.get(i).color();
        }
        if (style.radialGradient) {
            double radius = Math.max(1, Math.max(bounds.getWidth(), bounds.getHeight()) / 2d);
            return new RadialGradientPaint(
                    new Point2D.Double(bounds.getCenterX(), bounds.getCenterY()),
                    (float) radius, fractions, colors);
        }
        double angle = Math.toRadians(style.gradientAngle);
        double dx = Math.cos(angle);
        double dy = Math.sin(angle);
        double half = (Math.abs(bounds.getWidth() * dx) + Math.abs(bounds.getHeight() * dy)) / 2d;
        Point2D start = new Point2D.Double(bounds.getCenterX() - dx * half,
                bounds.getCenterY() - dy * half);
        Point2D end = new Point2D.Double(bounds.getCenterX() + dx * half,
                bounds.getCenterY() + dy * half);
        return new LinearGradientPaint(start, end, fractions, colors);
    }

    private static List<DrawingStyle.GradientStop> normalizeStops(
            List<DrawingStyle.GradientStop> stops) {
        List<DrawingStyle.GradientStop> result = new ArrayList<>();
        DrawingStyle.GradientStop first = stops.get(0);
        if (first.position() > 0.001) {
            result.add(new DrawingStyle.GradientStop(0, first.color()));
        }
        for (DrawingStyle.GradientStop stop : stops) {
            double pos = Units.clamp(stop.position(), 0, 1);
            if (!result.isEmpty()
                    && Math.abs(result.get(result.size() - 1).position() - pos) < 0.0001) {
                result.set(result.size() - 1, new DrawingStyle.GradientStop(pos, stop.color()));
            } else {
                result.add(new DrawingStyle.GradientStop(pos, stop.color()));
            }
        }
        DrawingStyle.GradientStop last = result.get(result.size() - 1);
        if (last.position() < 0.999) {
            result.add(new DrawingStyle.GradientStop(1, last.color()));
        }
        return result;
    }

    private static Color fillColor(DrawingStyle style) {
        if (style.gradientStops != null && !style.gradientStops.isEmpty()) {
            return style.gradientStops.get(0).color();
        }
        return style.fill == null ? Color.WHITE : style.fill;
    }

    private static Color fillModeColor(String mode, Color base) {
        return switch (mode == null ? "" : mode) {
            case "darkenLess" -> darkenLess(base);
            case "darken" -> darken(base, 0.6);
            case "lightenLess" -> lighten(base, 0.2);
            case "lighten" -> lighten(base, 0.4);
            default -> base;
        };
    }

    /**
     * ECMA-376 path fill "darkenLess", used for the folded-over corner flap.
     * Excel renders a white fill as #CDCDCD.
     */
    private static Color darkenLess(Color color) {
        double factor = 0.8;
        return new Color(
                (int) Math.round(color.getRed() * factor),
                (int) Math.round(color.getGreen() * factor),
                (int) Math.round(color.getBlue() * factor),
                color.getAlpha());
    }

    private static Color darken(Color color, double factor) {
        return new Color(
                (int) Math.round(color.getRed() * factor),
                (int) Math.round(color.getGreen() * factor),
                (int) Math.round(color.getBlue() * factor),
                color.getAlpha());
    }

    private static Color lighten(Color color, double amount) {
        return new Color(
                (int) Math.round(color.getRed() + (255 - color.getRed()) * amount),
                (int) Math.round(color.getGreen() + (255 - color.getGreen()) * amount),
                (int) Math.round(color.getBlue() + (255 - color.getBlue()) * amount),
                color.getAlpha());
    }

    public static void drawArrowHead(Graphics2D g2, double x, double y,
                                     double fromX, double fromY, String type, double lineWidth) {
        double angle = Math.atan2(y - fromY, x - fromX);
        // Excel keeps a minimum arrow size for hairlines but scales the head
        // with the line width for thick arrows.
        double size = Math.max(6, lineWidth * 4);
        double spread = 0.45;
        if ("arrow".equals(type)) {
            Path2D open = new Path2D.Double();
            open.moveTo(x - Math.cos(angle - spread) * size,
                    y - Math.sin(angle - spread) * size);
            open.lineTo(x, y);
            open.lineTo(x - Math.cos(angle + spread) * size,
                    y - Math.sin(angle + spread) * size);
            Stroke previous = g2.getStroke();
            float width = previous instanceof BasicStroke basic ? basic.getLineWidth() : 1f;
            g2.setStroke(new BasicStroke(width, BasicStroke.CAP_ROUND,
                    BasicStroke.JOIN_MITER));
            g2.draw(open);
            g2.setStroke(previous);
            return;
        }
        if ("diamond".equals(type)) {
            Path2D diamond = new Path2D.Double();
            diamond.moveTo(x, y);
            diamond.lineTo(x - Math.cos(angle) * size / 2 - Math.sin(angle) * size / 3,
                    y - Math.sin(angle) * size / 2 + Math.cos(angle) * size / 3);
            diamond.lineTo(x - Math.cos(angle) * size, y - Math.sin(angle) * size);
            diamond.lineTo(x - Math.cos(angle) * size / 2 + Math.sin(angle) * size / 3,
                    y - Math.sin(angle) * size / 2 - Math.cos(angle) * size / 3);
            diamond.closePath();
            g2.fill(diamond);
            return;
        }
        if ("oval".equals(type)) {
            double radius = size / 3;
            g2.fill(new Ellipse2D.Double(
                    x - Math.cos(angle) * radius - radius,
                    y - Math.sin(angle) * radius - radius, radius * 2, radius * 2));
            return;
        }
        Path2D path = new Path2D.Double();
        path.moveTo(x, y);
        path.lineTo(x - Math.cos(angle - spread) * size,
                y - Math.sin(angle - spread) * size);
        if ("stealth".equals(type)) {
            path.lineTo(x - Math.cos(angle) * size * 0.6,
                    y - Math.sin(angle) * size * 0.6);
        }
        path.lineTo(x - Math.cos(angle + spread) * size,
                y - Math.sin(angle + spread) * size);
        path.closePath();
        g2.fill(path);
    }

    public static void drawOutline(Graphics2D g2, Shape shape, DrawingStyle style) {
        if ("dbl".equalsIgnoreCase(style.compound)) {
            drawDoubleOutline(g2, shape, style);
            return;
        }
        g2.setStroke(stroke(style));
        g2.draw(shape);
    }

    /**
     * DrawingML cmpd="dbl" is one outline width divided into two equal lines
     * with a gap. Stroking a wide band and subtracting a narrower band leaves
     * the two parallel lines for arbitrary shape geometry.
     */
    private static void drawDoubleOutline(Graphics2D g2, Shape shape, DrawingStyle style) {
        double total = Math.max(0.75, style.lineWidth);
        double thickness = total / 3d;
        double gap = total - thickness * 2d;
        if (thickness < 0.15 || gap < 0.15) {
            g2.setStroke(stroke(style));
            g2.draw(shape);
            return;
        }
        Area outer = new Area(compoundStroke(style, total, total)
                .createStrokedShape(shape));
        Area inner = new Area(compoundStroke(style, gap, total)
                .createStrokedShape(shape));
        outer.subtract(inner);
        g2.fill(outer);
    }

    private static BasicStroke compoundStroke(DrawingStyle style, double width, double dashUnit) {
        float w = (float) Math.max(0.1, width);
        float unit = (float) Math.max(0.1, dashUnit);
        if ("dash".equals(style.dash) || "sysDash".equals(style.dash)) {
            return new BasicStroke(w, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER,
                    10f, new float[]{3f, 2f}, 0f);
        }
        if ("dot".equals(style.dash) || "sysDot".equals(style.dash)) {
            return new BasicStroke(w, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER,
                    10f, new float[]{1f, 2f}, 0f);
        }
        if ("lgDash".equals(style.dash)) {
            return new BasicStroke(w, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER,
                    10f, new float[]{unit * 8f, unit * 3f}, 0f);
        }
        return new BasicStroke(w, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER);
    }

    private static BasicStroke stroke(DrawingStyle style) {
        float width = (float) Math.max(0.25, style.lineWidth);
        if ("dash".equals(style.dash) || "sysDash".equals(style.dash)) {
            return new BasicStroke(width, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER,
                    10f, new float[]{3f, 2f}, 0f);
        }
        if ("dot".equals(style.dash) || "sysDot".equals(style.dash)) {
            return new BasicStroke(width, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER,
                    10f, new float[]{1f, 2f}, 0f);
        }
        if ("lgDash".equals(style.dash)) {
            return new BasicStroke(width, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER,
                    10f, new float[]{width * 8f, width * 3f}, 0f);
        }
        return new BasicStroke(width);
    }
}
