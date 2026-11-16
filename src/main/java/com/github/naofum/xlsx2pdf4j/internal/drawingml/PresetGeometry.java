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

import java.awt.BasicStroke;
import java.awt.Shape;
import java.awt.geom.AffineTransform;
import java.awt.geom.Arc2D;
import java.awt.geom.Area;
import java.awt.geom.Ellipse2D;
import java.awt.geom.Path2D;
import java.awt.geom.Point2D;
import java.awt.geom.Rectangle2D;
import java.awt.geom.RoundRectangle2D;

/**
 * Format-neutral ECMA-376 preset geometry. Host-format adapters supply the
 * shape frame and adjustment values; this class only constructs Java2D paths.
 */
public final class PresetGeometry {

    private PresetGeometry() {
    }

    public static Shape create(PresetShape preset, double x, double y, double w, double h,
                               ShapeAdjustments adj) {
        if (adj == null) {
            adj = ShapeAdjustments.NONE;
        }
        double arc = Math.min(w, h) * 0.15;
        return switch (preset) {
            case RECT, FLOW_CHART_PROCESS, FLOW_CHART_PREDEFINED_PROCESS ->
                    new Rectangle2D.Double(x, y, w, h);
            case ROUND_RECT, FLOW_CHART_ALTERNATE_PROCESS ->
                    new RoundRectangle2D.Double(x, y, w, h, arc, arc);
            case ELLIPSE, FLOW_CHART_CONNECTOR ->
                    new Ellipse2D.Double(x, y, w, h);
            case LINE, STRAIGHT_CONNECTOR_1 ->
                    new java.awt.geom.Line2D.Double(x, y, x + w, y + h);
            case LINE_INV ->
                    new java.awt.geom.Line2D.Double(x + w, y, x, y + h);
            case ARC -> arcShape(x, y, w, h);
            case TRIANGLE -> triangle(x, y, w, h);
            case RT_TRIANGLE -> rightTriangle(x, y, w, h);
            case DIAMOND -> diamond(x, y, w, h);
            case PARALLELOGRAM -> parallelogram(x, y, w, h);
            case TRAPEZOID -> trapezoid(x, y, w, h);
            case NON_ISOSCELES_TRAPEZOID ->
                    nonIsoscelesTrapezoid(x, y, w, h, adj);
            case PENTAGON -> regularPolygon(x, y, w, h, 5);
            case HEXAGON -> regularPolygon(x, y, w, h, 6);
            case HEPTAGON -> regularPolygon(x, y, w, h, 7);
            case OCTAGON -> octagon(x, y, w, h, adj);
            case DECAGON -> regularPolygon(x, y, w, h, 10);
            case DODECAGON -> regularPolygon(x, y, w, h, 12);
            case PIE -> pie(x, y, w, h);
            case PIE_WEDGE -> pieWedge(x, y, w, h);
            case BLOCK_ARC -> blockArc(x, y, w, h, adj);
            case DONUT -> donut(x, y, w, h);
            case NO_SMOKING -> noSmoking(x, y, w, h, adj);
            case CHORD -> chord(x, y, w, h);
            case CHEVRON -> chevron(x, y, w, h);
            case PLAQUE -> plaque(x, y, w, h, adj);
            case TEARDROP -> teardrop(x, y, w, h, adj);
            case HEART -> heart(x, y, w, h);
            case CLOUD -> cloud(x, y, w, h);
            case WAVE -> wave(x, y, w, h, adj);
            case DOUBLE_WAVE -> doubleWave(x, y, w, h, adj);
            case PLUS -> plus(x, y, w, h, adj);
            case MATH_PLUS -> mathPlus(x, y, w, h);
            case MATH_MINUS -> mathMinus(x, y, w, h);
            case MATH_MULTIPLY -> mathMultiply(x, y, w, h, adj);
            case MATH_DIVIDE -> mathDivide(x, y, w, h);
            case MATH_EQUAL -> mathEqual(x, y, w, h);
            case MATH_NOT_EQUAL -> mathNotEqual(x, y, w, h);
            case CHART_PLUS, CHART_X, CHART_STAR ->
                    new Rectangle2D.Double(x, y, w, h);
            case CORNER -> corner(x, y, w, h);
            case CORNER_TABS -> cornerTabs(x, y, w, h);
            case SQUARE_TABS -> squareTabs(x, y, w, h);
            case PLAQUE_TABS -> plaqueTabs(x, y, w, h);
            case FRAME -> frame(x, y, w, h);
            case HALF_FRAME -> halfFrame(x, y, w, h, adj);
            case BEVEL -> new Rectangle2D.Double(x, y, w, h);
            case CUBE -> cube(x, y, w, h, adj);
            case CAN -> can(x, y, w, h, adj);
            case FUNNEL -> funnel(x, y, w, h);
            case GEAR_6 -> gear(x, y, w, h, 6);
            case GEAR_9 -> gear(x, y, w, h, 9);
            case LIGHTNING_BOLT -> lightningBolt(x, y, w, h);
            case MOON -> moon(x, y, w, h);
            case SUN -> sun(x, y, w, h);
            case SMILEY_FACE -> new Ellipse2D.Double(x, y, w, h);
            case LEFT_BRACKET -> bracket(x, y, w, h, true);
            case RIGHT_BRACKET -> bracket(x, y, w, h, false);
            case BRACKET_PAIR -> bracketPair(x, y, w, h);
            case LEFT_BRACE -> brace(x, y, w, h, true);
            case RIGHT_BRACE -> brace(x, y, w, h, false);
            case BRACE_PAIR -> bracePair(x, y, w, h);
            case DIAG_STRIPE -> diagonalStripe(x, y, w, h, adj);
            case ROUND_1_RECT -> round1Rect(x, y, w, h, adj);
            case ROUND_2_SAME_RECT -> round2SameRect(x, y, w, h, adj);
            case ROUND_2_DIAG_RECT -> round2DiagRect(x, y, w, h, adj);
            case SNIP_1_RECT -> snip1Rect(x, y, w, h, adj);
            case SNIP_2_SAME_RECT -> snip2SameRect(x, y, w, h, adj);
            case SNIP_2_DIAG_RECT -> snip2DiagRect(x, y, w, h, adj);
            case SNIP_ROUND_RECT -> snipRoundRect(x, y, w, h, adj);
            case RIGHT_ARROW -> horizontalArrow(x, y, w, h, adj, false);
            case LEFT_ARROW -> horizontalArrow(x, y, w, h, adj, true);
            case UP_ARROW -> verticalArrow(x, y, w, h, adj, true);
            case DOWN_ARROW -> verticalArrow(x, y, w, h, adj, false);
            case NOTCHED_RIGHT_ARROW -> notchedRightArrow(x, y, w, h, adj);
            case STRIPED_RIGHT_ARROW -> stripedRightArrow(x, y, w, h, adj);
            case BENT_UP_ARROW -> bentUpArrow(x, y, w, h, adj);
            case UP_DOWN_ARROW -> upDownArrow(x, y, w, h, adj);
            case LEFT_UP_ARROW -> leftUpArrow(x, y, w, h, adj);
            case LEFT_RIGHT_UP_ARROW -> leftRightUpArrow(x, y, w, h, adj);
            case QUAD_ARROW -> quadArrow(x, y, w, h, adj);
            case BENT_ARROW -> bentArrow(x, y, w, h, adj);
            case UTURN_ARROW -> uturnArrow(x, y, w, h, adj);
            case SWOOSH_ARROW -> swooshArrow(x, y, w, h, adj);
            case CIRCULAR_ARROW -> circularArrow(x, y, w, h, false, false);
            case LEFT_CIRCULAR_ARROW -> circularArrow(x, y, w, h, true, false);
            case LEFT_RIGHT_CIRCULAR_ARROW -> circularArrow(x, y, w, h, false, true);
            case LEFT_RIGHT_ARROW ->
                    leftRightArrow(x, y, w, h, adj.value("adj1", 50000),
                            adj.value("adj2", 50000));
            case CURVED_LEFT_ARROW, CURVED_RIGHT_ARROW -> arrow(x, y, w, h, 0);
            case FLOW_CHART_DECISION -> diamond(x, y, w, h);
            case FLOW_CHART_DOCUMENT -> document(x, y, w, h);
            case FLOW_CHART_MULTIDOCUMENT -> multiDocumentFill(x, y, w, h);
            case FLOW_CHART_MAGNETIC_DISK -> diskFill(x, y, w, h);
            case FLOW_CHART_OFFPAGE_CONNECTOR -> offPageConnector(x, y, w, h);
            case FLOW_CHART_TERMINATOR -> flowChartTerminator(x, y, w, h);
            case FLOW_CHART_PREPARATION -> flowChartPreparation(x, y, w, h);
            case FLOW_CHART_MANUAL_INPUT -> flowChartManualInput(x, y, w, h);
            case FLOW_CHART_MANUAL_OPERATION -> flowChartManualOperation(x, y, w, h);
            case FLOW_CHART_PUNCHED_CARD -> flowChartPunchedCard(x, y, w, h);
            case FLOW_CHART_PUNCHED_TAPE -> flowChartPunchedTape(x, y, w, h);
            case FLOW_CHART_SUMMING_JUNCTION -> new Ellipse2D.Double(x, y, w, h);
            case FLOW_CHART_OR -> new Ellipse2D.Double(x, y, w, h);
            case FLOW_CHART_COLLATE -> flowChartCollate(x, y, w, h);
            case FLOW_CHART_SORT -> flowChartSort(x, y, w, h);
            case FLOW_CHART_EXTRACT -> flowChartExtract(x, y, w, h);
            case FLOW_CHART_MERGE -> flowChartMerge(x, y, w, h);
            case FLOW_CHART_OFFLINE_STORAGE -> flowChartOfflineStorage(x, y, w, h);
            case FLOW_CHART_ONLINE_STORAGE -> flowChartOnlineStorage(x, y, w, h);
            case FLOW_CHART_MAGNETIC_TAPE -> flowChartMagneticTape(x, y, w, h);
            case FLOW_CHART_MAGNETIC_DRUM -> flowChartMagneticDrum(x, y, w, h);
            case FLOW_CHART_DISPLAY -> flowChartDisplay(x, y, w, h);
            case FLOW_CHART_DELAY -> flowChartDelay(x, y, w, h);
            case FLOW_CHART_INPUT_OUTPUT -> flowChartInputOutput(x, y, w, h);
            case FLOW_CHART_INTERNAL_STORAGE -> new Rectangle2D.Double(x, y, w, h);
            case FOLDED_CORNER -> foldedCorner(x, y, w, h, adj.value("adj", 16667));
            case HOME_PLATE -> homePlate(x, y, w, h);
            case STAR_5 -> star(x, y, w, h, 5);
            case STAR_4 -> star(x, y, w, h, 4);
            case STAR_6 -> star(x, y, w, h, 6);
            case STAR_7 -> star(x, y, w, h, 7);
            case STAR_8 -> star(x, y, w, h, 8);
            case STAR_10 -> star(x, y, w, h, 10);
            case STAR_12 -> star(x, y, w, h, 12);
            case STAR_16 -> star(x, y, w, h, 16);
            case STAR_24 -> star(x, y, w, h, 24);
            case STAR_32 -> star(x, y, w, h, 32);
            case IRREGULAR_SEAL_1 -> irregularSeal1(x, y, w, h);
            case IRREGULAR_SEAL_2 -> star(x, y, w, h, 12);
            case BORDER_CALLOUT_1, BORDER_CALLOUT_2, BORDER_CALLOUT_3 ->
                    new Rectangle2D.Double(x, y, w, h);
            case CALLOUT_1, CALLOUT_2, CALLOUT_3,
                    ACCENT_CALLOUT_1, ACCENT_CALLOUT_2,
                    ACCENT_CALLOUT_3, ACCENT_BORDER_CALLOUT_1,
                    ACCENT_BORDER_CALLOUT_2, ACCENT_BORDER_CALLOUT_3 ->
                    new Rectangle2D.Double(x, y, w, h);
            case LEFT_ARROW_CALLOUT -> arrowCallout(x, y, w, h, adj, 1);
            case RIGHT_ARROW_CALLOUT -> arrowCallout(x, y, w, h, adj, 2);
            case UP_ARROW_CALLOUT -> arrowCallout(x, y, w, h, adj, 3);
            case DOWN_ARROW_CALLOUT -> arrowCallout(x, y, w, h, adj, 4);
            case LEFT_RIGHT_ARROW_CALLOUT -> arrowCallout(x, y, w, h, adj, 5);
            case UP_DOWN_ARROW_CALLOUT -> arrowCallout(x, y, w, h, adj, 6);
            case QUAD_ARROW_CALLOUT -> arrowCallout(x, y, w, h, adj, 7);
            case WEDGE_RECT_CALLOUT ->
                    wedgeRectCallout(x, y, w, h, adj.value("adj1", -20833),
                            adj.value("adj2", 62500));
            case WEDGE_ROUND_RECT_CALLOUT ->
                    callout(x, y, w, h, arc);
            case WEDGE_ELLIPSE_CALLOUT -> new Ellipse2D.Double(x, y, w, h);
            case CLOUD_CALLOUT -> new Ellipse2D.Double(x, y, w, h);
            case RIBBON -> ribbon(x, y, w, h, false);
            case RIBBON_2 -> ribbon(x, y, w, h, true);
            case ELLIPSE_RIBBON -> ellipseRibbon(x, y, w, h, false);
            case ELLIPSE_RIBBON_2 -> ellipseRibbon(x, y, w, h, true);
            case LEFT_RIGHT_RIBBON -> leftRightRibbon(x, y, w, h);
            case VERTICAL_SCROLL -> scroll(x, y, w, h, true);
            case HORIZONTAL_SCROLL -> scroll(x, y, w, h, false);
            default -> null;
        };
    }

    public static Shape arrow(double x, double y, double w, double h, double angle) {
        Path2D path = new Path2D.Double();
        double shaft = h * 0.28;
        double head = Math.min(w * 0.35, h * 0.5);
        path.moveTo(x, y + h / 2 - shaft / 2);
        path.lineTo(x + w - head, y + h / 2 - shaft / 2);
        path.lineTo(x + w - head, y);
        path.lineTo(x + w, y + h / 2);
        path.lineTo(x + w - head, y + h);
        path.lineTo(x + w - head, y + h / 2 + shaft / 2);
        path.lineTo(x, y + h / 2 + shaft / 2);
        path.closePath();
        if (angle == 0) {
            return path;
        }
        AffineTransform transform = AffineTransform.getRotateInstance(Math.toRadians(angle),
                x + w / 2, y + h / 2);
        return transform.createTransformedShape(path);
    }

    /**
     * ECMA-376 leftRightArrow geometry. adj1 controls the shaft width as a
     * percentage of the shape height; adj2 controls the arrowhead length
     * relative to the shorter shape side.
     */
    public static Shape leftRightArrow(double x, double y, double w, double h,
                                 double adj1, double adj2) {
        double ss = Math.max(0.001, Math.min(w, h));
        double maxAdj2 = 50000d * w / ss;
        double a1 = clamp(adj1, 0, 100000);
        double a2 = clamp(adj2, 0, maxAdj2);
        double x2 = ss * a2 / 100000d;
        double x3 = w - x2;
        double dy = h * a1 / 200000d;
        double y1 = h / 2d - dy;
        double y2 = h / 2d + dy;

        Path2D path = new Path2D.Double();
        path.moveTo(x, y + h / 2d);
        path.lineTo(x + x2, y);
        path.lineTo(x + x2, y + y1);
        path.lineTo(x + x3, y + y1);
        path.lineTo(x + x3, y);
        path.lineTo(x + w, y + h / 2d);
        path.lineTo(x + x3, y + h);
        path.lineTo(x + x3, y + y2);
        path.lineTo(x + x2, y + y2);
        path.lineTo(x + x2, y + h);
        path.closePath();
        return path;
    }

    public static Shape horizontalArrow(double x, double y, double w, double h,
                                  ShapeAdjustments adj, boolean left) {
        double adj1 = adj.value("adj1", 50000);
        double adj2 = adj.value("adj2", 50000);
        Shape result = rightArrow(x, y, w, h, adj1, adj2);
        return left ? flipHorizontal(result, x + w / 2d) : result;
    }

    public static Shape verticalArrow(double x, double y, double w, double h,
                                ShapeAdjustments adj, boolean up) {
        double adj1 = adj.value("adj1", 50000);
        double adj2 = adj.value("adj2", 50000);
        Shape result = downArrow(x, y, w, h, adj1, adj2);
        return up ? flipVertical(result, y + h / 2d) : result;
    }

    public static Shape rightArrow(double x, double y, double w, double h,
                             double adj1, double adj2) {
        double ss = Math.max(0.001, Math.min(w, h));
        double maxAdj2 = 100000d * w / ss;
        double a1 = clamp(adj1, 0, 100000);
        double a2 = clamp(adj2, 0, maxAdj2);
        double head = ss * a2 / 100000d;
        double x1 = w - head;
        double dy = h * a1 / 200000d;
        double y1 = h / 2d - dy;
        double y2 = h / 2d + dy;

        Path2D path = new Path2D.Double();
        path.moveTo(x, y + y1);
        path.lineTo(x + x1, y + y1);
        path.lineTo(x + x1, y);
        path.lineTo(x + w, y + h / 2d);
        path.lineTo(x + x1, y + h);
        path.lineTo(x + x1, y + y2);
        path.lineTo(x, y + y2);
        path.closePath();
        return path;
    }

    public static Shape downArrow(double x, double y, double w, double h,
                            double adj1, double adj2) {
        double ss = Math.max(0.001, Math.min(w, h));
        double maxAdj2 = 100000d * h / ss;
        double a1 = clamp(adj1, 0, 100000);
        double a2 = clamp(adj2, 0, maxAdj2);
        double head = ss * a2 / 100000d;
        double y1 = h - head;
        double dx = w * a1 / 200000d;
        double x1 = w / 2d - dx;
        double x2 = w / 2d + dx;

        Path2D path = new Path2D.Double();
        path.moveTo(x + x1, y);
        path.lineTo(x + x1, y + y1);
        path.lineTo(x, y + y1);
        path.lineTo(x + w / 2d, y + h);
        path.lineTo(x + w, y + y1);
        path.lineTo(x + x2, y + y1);
        path.lineTo(x + x2, y);
        path.closePath();
        return path;
    }

    public static Shape notchedRightArrow(double x, double y, double w, double h,
                                    ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double a1 = clamp(adj.value("adj1", 50000), 0, 100000);
        double maxAdj2 = 100000d * w / ss;
        double a2 = clamp(adj.value("adj2", 50000), 0, maxAdj2);
        double dx2 = ss * a2 / 100000d;
        double x2 = w - dx2;
        double dy1 = h * a1 / 200000d;
        double y1 = h / 2d - dy1;
        double y2 = h / 2d + dy1;
        double x1 = dy1 * dx2 / (h / 2d);

        Path2D path = new Path2D.Double();
        path.moveTo(x, y + y1);
        path.lineTo(x + x2, y + y1);
        path.lineTo(x + x2, y);
        path.lineTo(x + w, y + h / 2d);
        path.lineTo(x + x2, y + h);
        path.lineTo(x + x2, y + y2);
        path.lineTo(x, y + y2);
        path.lineTo(x + x1, y + h / 2d);
        path.closePath();
        return path;
    }

    public static Shape stripedRightArrow(double x, double y, double w, double h,
                                    ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double a1 = clamp(adj.value("adj1", 50000), 0, 100000);
        double maxAdj2 = 84375d * w / ss;
        double a2 = clamp(adj.value("adj2", 50000), 0, maxAdj2);
        double x4 = ss * 5d / 32d;
        double x5 = w - ss * a2 / 100000d;
        double dy1 = h * a1 / 200000d;
        double y1 = h / 2d - dy1;
        double y2 = h / 2d + dy1;
        double x6 = w - dy1 * (w - x5) / (h / 2d);

        Path2D path = new Path2D.Double();
        // Two tail stripes, then the arrow body. The gaps are the white
        // spaces between the three subpaths.
        tailStripe(path, x, y, ss / 32d, y1, y2);
        tailStripe(path, x + ss / 16d, y, ss / 8d, y1, y2);
        path.moveTo(x + x4, y + y1);
        path.lineTo(x + x5, y + y1);
        path.lineTo(x + x5, y);
        path.lineTo(x + w, y + h / 2d);
        path.lineTo(x + x5, y + h);
        path.lineTo(x + x5, y + y2);
        path.lineTo(x + x4, y + y2);
        path.closePath();
        return path;
    }

    public static void tailStripe(Path2D path, double x, double y,
                                   double width, double y1, double y2) {
        path.moveTo(x, y + y1);
        path.lineTo(x + width, y + y1);
        path.lineTo(x + width, y + y2);
        path.lineTo(x, y + y2);
        path.closePath();
    }

    public static Shape bentUpArrow(double x, double y, double w, double h,
                              ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double a1 = clamp(adj.value("adj1", 25000), 0, 50000);
        double a2 = clamp(adj.value("adj2", 25000), 0, 50000);
        double a3 = clamp(adj.value("adj3", 25000), 0, 50000);
        double y1 = ss * a3 / 100000d;
        double x1 = w - ss * a2 / 50000d;
        double x3 = w - ss * a2 / 100000d;
        double dx2 = ss * a1 / 200000d;
        double x2 = x3 - dx2;
        double x4 = x3 + dx2;
        double y2 = h - ss * a1 / 100000d;

        Path2D path = new Path2D.Double();
        path.moveTo(x, y + y2);
        path.lineTo(x + x2, y + y2);
        path.lineTo(x + x2, y + y1);
        path.lineTo(x + x1, y + y1);
        path.lineTo(x + x3, y);
        path.lineTo(x + w, y + y1);
        path.lineTo(x + x4, y + y1);
        path.lineTo(x + x4, y + h);
        path.lineTo(x, y + h);
        path.closePath();
        return path;
    }

    public static Shape upDownArrow(double x, double y, double w, double h,
                              ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double a1 = clamp(adj.value("adj1", 50000), 0, 100000);
        double maxAdj2 = 50000d * h / ss;
        double a2 = clamp(adj.value("adj2", 50000), 0, maxAdj2);
        double y2 = ss * a2 / 100000d;
        double y3 = h - y2;
        double dx1 = w * a1 / 200000d;
        double x1 = w / 2d - dx1;
        double x2 = w / 2d + dx1;
        double dy1 = x1 * y2 / (w / 2d);

        Path2D path = new Path2D.Double();
        path.moveTo(x, y + y2);
        path.lineTo(x + w / 2d, y);
        path.lineTo(x + w, y + y2);
        path.lineTo(x + x2, y + y2);
        path.lineTo(x + x2, y + y3);
        path.lineTo(x + w, y + y3);
        path.lineTo(x + w / 2d, y + h);
        path.lineTo(x, y + y3);
        path.lineTo(x + x1, y + y3);
        path.lineTo(x + x1, y + y2);
        path.closePath();
        return path;
    }

    public static Shape leftUpArrow(double x, double y, double w, double h,
                              ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double a2 = clamp(adj.value("adj2", 25000), 0, 50000);
        double a1 = clamp(adj.value("adj1", 25000), 0, a2 * 2d);
        double a3 = clamp(adj.value("adj3", 25000), 0,
                100000d - a2 * 2d);
        double x1 = ss * a3 / 100000d;
        double dx2 = ss * a2 / 50000d;
        double x2 = w - dx2;
        double y2 = h - dx2;
        double x4 = w - ss * a2 / 100000d;
        double y4 = h - ss * a2 / 100000d;
        double dx3 = ss * a1 / 200000d;
        double x3 = x4 - dx3;
        double x5 = x4 + dx3;
        double y3 = y4 - dx3;
        double y5 = y4 + dx3;

        Path2D path = new Path2D.Double();
        path.moveTo(x, y + y4);
        path.lineTo(x + x1, y + y2);
        path.lineTo(x + x1, y + y3);
        path.lineTo(x + x3, y + y3);
        path.lineTo(x + x3, y + x1);
        path.lineTo(x + x2, y + x1);
        path.lineTo(x + x4, y);
        path.lineTo(x + w, y + x1);
        path.lineTo(x + x5, y + x1);
        path.lineTo(x + x5, y + y5);
        path.lineTo(x + x1, y + y5);
        path.lineTo(x + x1, y + h);
        path.closePath();
        return path;
    }

    public static Shape leftRightUpArrow(double x, double y, double w, double h,
                                   ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double a2 = clamp(adj.value("adj2", 25000), 0, 50000);
        double a1 = clamp(adj.value("adj1", 25000), 0, a2 * 2d);
        double maxAdj3 = (100000d - a2 * 2d) / 2d;
        double a3 = clamp(adj.value("adj3", 25000), 0, maxAdj3);
        double x1 = ss * a3 / 100000d;
        double dx2 = ss * a2 / 100000d;
        double x2 = w / 2d - dx2;
        double x5 = w / 2d + dx2;
        double dx3 = ss * a1 / 200000d;
        double x3 = w / 2d - dx3;
        double x4 = w / 2d + dx3;
        double x6 = w - x1;
        double dy2 = ss * a2 / 50000d;
        double y2 = h - dy2;
        double y4 = h - dx2;
        double y3 = y4 - dx3;
        double y5 = y4 + dx3;

        Path2D path = new Path2D.Double();
        path.moveTo(x, y + y4);
        path.lineTo(x + x1, y + y2);
        path.lineTo(x + x1, y + y3);
        path.lineTo(x + x3, y + y3);
        path.lineTo(x + x3, y + x1);
        path.lineTo(x + x2, y + x1);
        path.lineTo(x + w / 2d, y);
        path.lineTo(x + x5, y + x1);
        path.lineTo(x + x4, y + x1);
        path.lineTo(x + x4, y + y3);
        path.lineTo(x + x6, y + y3);
        path.lineTo(x + x6, y + y2);
        path.lineTo(x + w, y + y4);
        path.lineTo(x + x6, y + h);
        path.lineTo(x + x6, y + y5);
        path.lineTo(x + x1, y + y5);
        path.lineTo(x + x1, y + h);
        path.closePath();
        return path;
    }

    public static Shape quadArrow(double x, double y, double w, double h,
                            ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double a2 = clamp(adj.value("adj2", 22500), 0, 50000);
        double a1 = clamp(adj.value("adj1", 22500), 0, a2 * 2d);
        double maxAdj3 = (100000d - a2 * 2d) / 2d;
        double a3 = clamp(adj.value("adj3", 22500), 0, maxAdj3);
        double x1 = ss * a3 / 100000d;
        double dx2 = ss * a2 / 100000d;
        double x2 = w / 2d - dx2;
        double x5 = w / 2d + dx2;
        double dx3 = ss * a1 / 200000d;
        double x3 = w / 2d - dx3;
        double x4 = w / 2d + dx3;
        double x6 = w - x1;
        double y2 = h / 2d - dx2;
        double y5 = h / 2d + dx2;
        double y3 = h / 2d - dx3;
        double y4 = h / 2d + dx3;
        double y6 = h - x1;

        Path2D path = new Path2D.Double();
        path.moveTo(x, y + h / 2d);
        path.lineTo(x + x1, y + y2);
        path.lineTo(x + x1, y + y3);
        path.lineTo(x + x3, y + y3);
        path.lineTo(x + x3, y + x1);
        path.lineTo(x + x2, y + x1);
        path.lineTo(x + w / 2d, y);
        path.lineTo(x + x5, y + x1);
        path.lineTo(x + x4, y + x1);
        path.lineTo(x + x4, y + y3);
        path.lineTo(x + x6, y + y3);
        path.lineTo(x + x6, y + y2);
        path.lineTo(x + w, y + h / 2d);
        path.lineTo(x + x6, y + y5);
        path.lineTo(x + x6, y + y4);
        path.lineTo(x + x4, y + y4);
        path.lineTo(x + x4, y + y6);
        path.lineTo(x + x5, y + y6);
        path.lineTo(x + w / 2d, y + h);
        path.lineTo(x + x2, y + y6);
        path.lineTo(x + x3, y + y6);
        path.lineTo(x + x3, y + y4);
        path.lineTo(x + x1, y + y4);
        path.lineTo(x + x1, y + y5);
        path.closePath();
        return path;
    }

    public static Shape bentArrow(double x, double y, double w, double h,
                            ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double a2 = clamp(adj.value("adj2", 25000), 0, 50000);
        double a1 = clamp(adj.value("adj1", 25000), 0, a2 * 2d);
        double a3 = clamp(adj.value("adj3", 25000), 0, 50000);
        double th = ss * a1 / 100000d;
        double aw2 = ss * a2 / 100000d;
        double th2 = th / 2d;
        double dh2 = aw2 - th2;
        double ah = ss * a3 / 100000d;
        double bw = w - ah;
        double bh = h - dh2;
        double bs = Math.min(bw, bh);
        double maxAdj4 = 100000d * bs / ss;
        double a4 = clamp(adj.value("adj4", 43750), 0, maxAdj4);
        double bd = ss * a4 / 100000d;
        double bd2 = Math.max(bd - th, 0);
        double x3 = th + bd2;
        double x4 = w - ah;
        double y3 = dh2 + th;
        double y4 = y3 + dh2;
        double y5 = dh2 + bd;

        Path2D path = new Path2D.Double();
        path.moveTo(x, y + h);
        path.lineTo(x, y + y5);
        arcTo(path, bd, bd, 180, 90);
        path.lineTo(x + x4, y + dh2);
        path.lineTo(x + x4, y);
        path.lineTo(x + w, y + aw2);
        path.lineTo(x + x4, y + y4);
        path.lineTo(x + x4, y + y3);
        path.lineTo(x + x3, y + y3);
        arcTo(path, bd2, bd2, 270, -90);
        path.lineTo(x + th, y + h);
        path.closePath();
        return path;
    }

    public static Shape uturnArrow(double x, double y, double w, double h,
                             ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double a2 = clamp(adj.value("adj2", 25000), 0, 25000);
        double a1 = clamp(adj.value("adj1", 25000), 0, a2 * 2d);
        double q2 = a1 * ss / h;
        double q3 = 100000d - q2;
        double maxAdj3 = q3 * h / ss;
        double a3 = clamp(adj.value("adj3", 25000), 0, maxAdj3);
        double minAdj5 = (a3 - a1) * ss / h;
        double a5 = clamp(adj.value("adj5", 75000), minAdj5, 100000);
        double th = ss * a1 / 100000d;
        double aw2 = ss * a2 / 100000d;
        double dh2 = aw2 - th / 2d;
        double y5 = h * a5 / 100000d;
        double ah = ss * a3 / 100000d;
        double y4 = y5 - ah;
        double x9 = w - dh2;
        double bs = Math.min(x9 / 2d, y4);
        double maxAdj4 = 100000d * bs / ss;
        double a4 = clamp(adj.value("adj4", 43750), 0, maxAdj4);
        double bd = ss * a4 / 100000d;
        double bd2 = Math.max(bd - th, 0);
        double x3 = th + bd2;
        double x8 = w - aw2;
        double x6 = x8 - aw2;
        double x7 = x6 + dh2;
        double x4 = x9 - bd;

        Path2D path = new Path2D.Double();
        path.moveTo(x, y + h);
        path.lineTo(x, y + bd);
        arcTo(path, bd, bd, 180, 90);
        path.lineTo(x + x4, y);
        arcTo(path, bd, bd, 270, 90);
        path.lineTo(x + x9, y + y4);
        path.lineTo(x + w, y + y4);
        path.lineTo(x + x8, y + y5);
        path.lineTo(x + x6, y + y4);
        path.lineTo(x + x7, y + y4);
        path.lineTo(x + x7, y + x3);
        arcTo(path, bd2, bd2, 0, -90);
        path.lineTo(x + x3, y + th);
        arcTo(path, bd2, bd2, 270, -90);
        path.lineTo(x + th, y + h);
        path.closePath();
        return path;
    }

    public static Shape swooshArrow(double x, double y, double w, double h,
                              ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double a1 = clamp(adj.value("adj1", 25000), 1, 75000);
        double maxAdj2 = 70000d * w / ss;
        double a2 = clamp(adj.value("adj2", 16667), 0, maxAdj2);
        double ad1 = h * a1 / 100000d;
        double ad2 = ss * a2 / 100000d;
        double xB = w - ad2;
        double yB = ss / 8d;
        double alfa = Math.toRadians(90d / 14d);
        double dx0 = ss / 8d * Math.tan(alfa);
        double xC = xB - dx0;
        double dx1 = ad1 * Math.tan(alfa);
        double yF = yB + ad1;
        double xF = xB + dx1;
        double xE = xF + dx0;
        double yE = yF + ss / 8d;
        double yD = yE / 2d - h / 20d;
        double yP2 = yF + h / 12d;

        Path2D path = new Path2D.Double();
        path.moveTo(x, y + h);
        path.quadTo(x + w / 6d, y + h / 3d, x + xB, y + yB);
        path.lineTo(x + xC, y);
        path.lineTo(x + w, y + yD);
        path.lineTo(x + xE, y + yE);
        path.lineTo(x + xF, y + yF);
        path.quadTo(x + w / 4d, y + yP2, x, y + h);
        path.closePath();
        return path;
    }

    public static Shape circularArrow(double x, double y, double w, double h,
                                boolean left, boolean doubleHeaded) {        double cx = x + w / 2d;
        double cy = y + h / 2d;
        double rx = Math.max(0.1, w * 0.42);
        double ry = Math.max(0.1, h * 0.42);
        double band = Math.max(0.8, Math.min(w, h) * 0.09);
        double start = Math.toRadians(-60);
        double sweep = Math.toRadians(doubleHeaded ? 300 : 290);
        double end = start + sweep;

        Area area = new Area(ringBand(cx, cy, rx, ry, band, start, end));
        area.add(new Area(arrowHead(cx, cy, rx, ry, band, end, 1)));
        if (doubleHeaded) {
            area.add(new Area(arrowHead(cx, cy, rx, ry, band, start, -1)));
        }
        if (left) {
            AffineTransform mirror = new AffineTransform(-1, 0, 0, 1, cx * 2d, 0);
            area.transform(mirror);
        }
        return area;
    }

    public static Shape ringBand(double cx, double cy, double rx, double ry,
                                  double band, double start, double end) {
        int steps = 72;
        double innerRx = Math.max(0.1, rx - band);
        double innerRy = Math.max(0.1, ry - band);
        Path2D path = new Path2D.Double();
        for (int i = 0; i <= steps; i++) {
            double t = start + (end - start) * i / steps;
            double px = cx + rx * Math.cos(t);
            double py = cy + ry * Math.sin(t);
            if (i == 0) {
                path.moveTo(px, py);
            } else {
                path.lineTo(px, py);
            }
        }
        for (int i = steps; i >= 0; i--) {
            double t = start + (end - start) * i / steps;
            path.lineTo(cx + innerRx * Math.cos(t),
                    cy + innerRy * Math.sin(t));
        }
        path.closePath();
        return path;
    }

    public static Shape arrowHead(double cx, double cy, double rx, double ry,
                                   double band, double angle, int direction) {
        double midRx = rx - band / 2d;
        double midRy = ry - band / 2d;
        double baseX = cx + midRx * Math.cos(angle);
        double baseY = cy + midRy * Math.sin(angle);
        double tangentX = -rx * Math.sin(angle);
        double tangentY = ry * Math.cos(angle);
        double length = Math.hypot(tangentX, tangentY);
        if (length == 0) {
            length = 1;
        }
        tangentX /= length;
        tangentY /= length;
        double perpendicularX = -tangentY;
        double perpendicularY = tangentX;
        double half = band * 0.95;
        double headLength = band * 1.8;
        Path2D path = new Path2D.Double();
        path.moveTo(baseX + perpendicularX * half
                        + tangentX * headLength * direction,
                baseY + perpendicularY * half
                        + tangentY * headLength * direction);
        path.lineTo(baseX + perpendicularX * half
                        - tangentX * headLength * direction,
                baseY + perpendicularY * half
                        - tangentY * headLength * direction);
        path.lineTo(baseX - perpendicularX * half
                        - tangentX * headLength * direction,
                baseY - perpendicularY * half
                        - tangentY * headLength * direction);
        path.closePath();
        return path;
    }

    public static void arcTo(Path2D path, double rx, double ry,
                              double startDegrees, double sweepDegrees) {
        Point2D current = path.getCurrentPoint();
        if (current == null || rx <= 0 || ry <= 0) {
            return;
        }
        double angle = Math.toRadians(startDegrees);
        double centerX = current.getX() - rx * Math.cos(angle);
        double centerY = current.getY() - ry * Math.sin(angle);
        path.append(new Arc2D.Double(centerX - rx, centerY - ry,
                2 * rx, 2 * ry, -startDegrees, -sweepDegrees, Arc2D.OPEN), true);
    }

    public static Shape flipHorizontal(Shape shape, double centerX) {
        AffineTransform transform = new AffineTransform();
        transform.translate(centerX * 2d, 0);
        transform.scale(-1, 1);
        return transform.createTransformedShape(shape);
    }

    public static Shape flipVertical(Shape shape, double centerY) {
        AffineTransform transform = new AffineTransform();
        transform.translate(0, centerY * 2d);
        transform.scale(1, -1);
        return transform.createTransformedShape(shape);
    }

    public static Shape diamond(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x + w / 2, y);
        path.lineTo(x + w, y + h / 2);
        path.lineTo(x + w / 2, y + h);
        path.lineTo(x, y + h / 2);
        path.closePath();
        return path;
    }

    /**
     * Exact ECMA-376 flowChartDocument outline: a single cubic Bezier bottom
     * edge with guides 17322/20172/23922 of 21600.
     */
    public static Shape document(double x, double y, double w, double h) {
        double y1 = y + h * 17322 / 21600.0;
        double y2 = y + h * 20172 / 21600.0;
        double controlY = y + h * 23922 / 21600.0;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y);
        path.lineTo(x + w, y);
        path.lineTo(x + w, y1);
        path.curveTo(x + w / 2, y1, x + w / 2, controlY, x, y2);
        path.closePath();
        return path;
    }

    /**
     * ECMA-376 flowChartMultidocument fill: three offset document sheets. The
     * front sheet is bottom-left; the back sheets peek out at the top-right.
     */
    public static Shape multiDocumentFill(double x, double y, double w, double h) {
        double sx = w / 21600.0;
        double sy = h / 21600.0;
        Path2D path = new Path2D.Double();
        // Front/bottom-left sheet.
        path.moveTo(x, y + 20782 * sy);
        path.curveTo(x + 9298 * sx, y + 23542 * sy,
                x + 9298 * sx, y + 18022 * sy,
                x + 18595 * sx, y + 18022 * sy);
        path.lineTo(x + 18595 * sx, y + 3675 * sy);
        path.lineTo(x, y + 3675 * sy);
        path.closePath();
        // Middle sheet.
        path.moveTo(x + 1532 * sx, y + 3675 * sy);
        path.lineTo(x + 1532 * sx, y + 1815 * sy);
        path.lineTo(x + 20000 * sx, y + 1815 * sy);
        path.lineTo(x + 20000 * sx, y + 16252 * sy);
        path.curveTo(x + 19298 * sx, y + 16252 * sy,
                x + 18595 * sx, y + 16352 * sy,
                x + 18595 * sx, y + 16352 * sy);
        path.lineTo(x + 18595 * sx, y + 3675 * sy);
        path.closePath();
        // Back/top-right sheet.
        path.moveTo(x + 2972 * sx, y + 1815 * sy);
        path.lineTo(x + 2972 * sx, y);
        path.lineTo(x + 21600 * sx, y);
        path.lineTo(x + 21600 * sx, y + 14392 * sy);
        path.curveTo(x + 20800 * sx, y + 14392 * sy,
                x + 20000 * sx, y + 14467 * sy,
                x + 20000 * sx, y + 14467 * sy);
        path.lineTo(x + 20000 * sx, y + 1815 * sy);
        path.closePath();
        return path;
    }

    /**
     * The stroke path omits the parts hidden behind the front sheets.
     */
    public static Shape multiDocumentOutline(double x, double y, double w, double h) {
        double sx = w / 21600.0;
        double sy = h / 21600.0;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y + 3675 * sy);
        path.lineTo(x + 18595 * sx, y + 3675 * sy);
        path.lineTo(x + 18595 * sx, y + 18022 * sy);
        path.curveTo(x + 9298 * sx, y + 18022 * sy,
                x + 9298 * sx, y + 23542 * sy,
                x, y + 20782 * sy);
        path.closePath();
        path.moveTo(x + 1532 * sx, y + 3675 * sy);
        path.lineTo(x + 1532 * sx, y + 1815 * sy);
        path.lineTo(x + 20000 * sx, y + 1815 * sy);
        path.lineTo(x + 20000 * sx, y + 16252 * sy);
        path.curveTo(x + 19298 * sx, y + 16252 * sy,
                x + 18595 * sx, y + 16352 * sy,
                x + 18595 * sx, y + 16352 * sy);
        path.moveTo(x + 2972 * sx, y + 1815 * sy);
        path.lineTo(x + 2972 * sx, y);
        path.lineTo(x + 21600 * sx, y);
        path.lineTo(x + 21600 * sx, y + 14392 * sy);
        path.curveTo(x + 20800 * sx, y + 14392 * sy,
                x + 20000 * sx, y + 14467 * sy,
                x + 20000 * sx, y + 14467 * sy);
        return path;
    }

    /**
     * ECMA-376 flowChartMagneticDisk. The fill is the upper half of the top
     * ellipse + body + lower half of the bottom ellipse. Only the front arcs
     * are stroked; the bottom back arc and the horizontal diameters are not.
     */
    public static Shape diskFill(double x, double y, double w, double h) {
        double ry = h / 6.0;
        double topCy = y + ry;
        double bottomCy = y + h - ry;
        Path2D path = new Path2D.Double();
        path.moveTo(x, topCy);
        // Left -> top -> right.
        path.append(new Arc2D.Double(x, y, w, 2 * ry, 180, -180, Arc2D.OPEN), true);
        path.lineTo(x + w, bottomCy);
        // Right -> bottom -> left (front arc only).
        path.append(new Arc2D.Double(x, y + h - 2 * ry, w, 2 * ry, 0, -180, Arc2D.OPEN), true);
        path.closePath();
        return path;
    }

    public static Shape diskOutline(double x, double y, double w, double h) {
        double ry = h / 6.0;
        double topCy = y + ry;
        double bottomCy = y + h - ry;
        Path2D path = new Path2D.Double();
        // Outer silhouette: top back arc, sides, bottom front arc.
        path.moveTo(x, topCy);
        path.append(new Arc2D.Double(x, y, w, 2 * ry, 180, -180, Arc2D.OPEN), true);
        path.lineTo(x + w, bottomCy);
        path.append(new Arc2D.Double(x, y + h - 2 * ry, w, 2 * ry, 0, -180, Arc2D.OPEN), true);
        path.closePath();
        // Front rim of the top ellipse.
        path.moveTo(x + w, topCy);
        path.append(new Arc2D.Double(x, y, w, 2 * ry, 0, -180, Arc2D.OPEN), true);
        return path;
    }

    /**
     * Flowchart off-page connector: a pentagon with a flat top and a point at
     * the bottom. It must not use the right-pointing home-plate geometry.
     */
    public static Shape offPageConnector(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        double tip = Math.min(h * 0.2, w * 0.5);
        path.moveTo(x, y);
        path.lineTo(x + w, y);
        path.lineTo(x + w, y + h - tip);
        path.lineTo(x + w / 2, y + h);
        path.lineTo(x, y + h - tip);
        path.closePath();
        return path;
    }

    public static Shape homePlate(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        double tip = Math.min(w * 0.2, h * 0.5);
        path.moveTo(x, y);
        path.lineTo(x + w - tip, y);
        path.lineTo(x + w, y + h / 2);
        path.lineTo(x + w - tip, y + h);
        path.lineTo(x, y + h);
        path.closePath();
        return path;
    }

    /**
     * ECMA-376 foldedCorner main fill polygon (rectangle with the bottom-right
     * corner cut off). The flap fill and crease are drawn separately.
     */
    public static Shape flowChartTerminator(double x, double y, double w, double h) {
        double rx = w * 3475d / 21600d;
        double ry = h / 2d;
        Path2D path = new Path2D.Double();
        path.moveTo(x + rx, y);
        path.lineTo(x + w - rx, y);
        arcTo(path, rx, ry, 270, 180);
        path.lineTo(x + rx, y + h);
        arcTo(path, rx, ry, 90, 180);
        path.closePath();
        return path;
    }

    public static Shape flowChartPreparation(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x, y + h / 2d);
        path.lineTo(x + w / 5d, y);
        path.lineTo(x + w * 4d / 5d, y);
        path.lineTo(x + w, y + h / 2d);
        path.lineTo(x + w * 4d / 5d, y + h);
        path.lineTo(x + w / 5d, y + h);
        path.closePath();
        return path;
    }

    public static Shape flowChartManualInput(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x, y + h / 5d);
        path.lineTo(x + w, y);
        path.lineTo(x + w, y + h);
        path.lineTo(x, y + h);
        path.closePath();
        return path;
    }

    public static Shape flowChartManualOperation(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x, y);
        path.lineTo(x + w, y);
        path.lineTo(x + w * 4d / 5d, y + h);
        path.lineTo(x + w / 5d, y + h);
        path.closePath();
        return path;
    }

    public static Shape flowChartPunchedCard(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x, y + h / 5d);
        path.lineTo(x + w / 5d, y);
        path.lineTo(x + w, y);
        path.lineTo(x + w, y + h);
        path.lineTo(x, y + h);
        path.closePath();
        return path;
    }

    public static Shape flowChartPunchedTape(double x, double y, double w, double h) {
        double rx = w / 4d;
        double ry = h / 10d;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y + ry);
        arcTo(path, rx, ry, 180, -180);
        arcTo(path, rx, ry, 180, 180);
        path.lineTo(x + w, y + h - ry);
        arcTo(path, rx, ry, 0, -180);
        arcTo(path, rx, ry, 0, 180);
        path.closePath();
        return path;
    }

    public static Shape flowChartCollate(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x, y);
        path.lineTo(x + w, y);
        path.lineTo(x + w / 2d, y + h / 2d);
        path.lineTo(x + w, y + h);
        path.lineTo(x, y + h);
        path.lineTo(x + w / 2d, y + h / 2d);
        path.closePath();
        return path;
    }

    public static Shape flowChartSort(double x, double y, double w, double h) {
        return diamond(x, y, w, h);
    }

    public static Shape flowChartExtract(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x, y + h);
        path.lineTo(x + w / 2d, y);
        path.lineTo(x + w, y + h);
        path.closePath();
        return path;
    }

    public static Shape flowChartMerge(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x, y);
        path.lineTo(x + w, y);
        path.lineTo(x + w / 2d, y + h);
        path.closePath();
        return path;
    }

    public static Shape flowChartOfflineStorage(double x, double y, double w, double h) {
        return flowChartMerge(x, y, w, h);
    }

    public static Shape flowChartOnlineStorage(double x, double y, double w, double h) {
        double rx = w / 6d;
        double ry = h / 2d;
        Path2D path = new Path2D.Double();
        path.moveTo(x + rx, y);
        path.lineTo(x + w, y);
        arcTo(path, rx, ry, 270, -180);
        path.lineTo(x + rx, y + h);
        arcTo(path, rx, ry, 90, 180);
        path.closePath();
        return path;
    }

    public static Shape flowChartMagneticTape(double x, double y, double w, double h) {
        double rx = w / 2d;
        double ry = h / 2d;
        double angle = Math.toDegrees(Math.atan2(h, w));
        double ib = y + h / 2d + h / 2d * Math.sin(Math.toRadians(45));
        Path2D path = new Path2D.Double();
        path.moveTo(x + w / 2d, y + h);
        arcTo(path, rx, ry, 90, 90);
        arcTo(path, rx, ry, 180, 90);
        arcTo(path, rx, ry, 270, 90);
        arcTo(path, rx, ry, 0, angle);
        path.lineTo(x + w, ib);
        path.lineTo(x + w, y + h);
        path.closePath();
        return path;
    }

    public static Shape flowChartMagneticDrum(double x, double y, double w, double h) {
        double rx = w / 6d;
        double ry = h / 2d;
        Path2D path = new Path2D.Double();
        path.moveTo(x + rx, y);
        path.lineTo(x + w - rx, y);
        arcTo(path, rx, ry, 270, 180);
        path.lineTo(x + rx, y + h);
        arcTo(path, rx, ry, 90, 180);
        path.closePath();
        return path;
    }

    public static Shape flowChartDisplay(double x, double y, double w, double h) {
        double rx = w / 6d;
        double ry = h / 2d;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y + h / 2d);
        path.lineTo(x + rx, y);
        path.lineTo(x + w - rx, y);
        arcTo(path, rx, ry, 270, 180);
        path.lineTo(x + rx, y + h);
        path.closePath();
        return path;
    }

    public static Shape flowChartDelay(double x, double y, double w, double h) {
        double rx = w / 2d;
        double ry = h / 2d;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y);
        path.lineTo(x + w / 2d, y);
        arcTo(path, rx, ry, 270, 180);
        path.lineTo(x, y + h);
        path.closePath();
        return path;
    }

    public static Shape flowChartInputOutput(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x, y + h);
        path.lineTo(x + w / 5d, y);
        path.lineTo(x + w, y);
        path.lineTo(x + w * 4d / 5d, y + h);
        path.closePath();
        return path;
    }

    public static Shape round1Rect(double x, double y, double w, double h, ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double r = ss * clamp(adj.value("adj", 16667), 0, 50000) / 100000d;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y);
        path.lineTo(x + w - r, y);
        if (r > 0.01) {
            arcTo(path, r, r, 270, 90);
        } else {
            path.lineTo(x + w, y);
        }
        path.lineTo(x + w, y + h);
        path.lineTo(x, y + h);
        path.closePath();
        return path;
    }

    public static Shape round2SameRect(double x, double y, double w, double h, ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double r1 = ss * clamp(adj.value("adj1", 16667), 0, 50000) / 100000d;
        double r2 = ss * clamp(adj.value("adj2", 0), 0, 50000) / 100000d;
        Path2D path = new Path2D.Double();
        path.moveTo(x + r1, y);
        path.lineTo(x + w - r1, y);
        if (r1 > 0.01) {
            arcTo(path, r1, r1, 270, 90);
        } else {
            path.lineTo(x + w, y);
        }
        path.lineTo(x + w, y + h - r2);
        if (r2 > 0.01) {
            arcTo(path, r2, r2, 0, 90);
        } else {
            path.lineTo(x + w, y + h);
        }
        path.lineTo(x + r2, y + h);
        if (r2 > 0.01) {
            arcTo(path, r2, r2, 90, 90);
        } else {
            path.lineTo(x, y + h);
        }
        path.lineTo(x, y + r1);
        if (r1 > 0.01) {
            arcTo(path, r1, r1, 180, 90);
        } else {
            path.lineTo(x, y);
        }
        path.closePath();
        return path;
    }

    public static Shape round2DiagRect(double x, double y, double w, double h, ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double r1 = ss * clamp(adj.value("adj1", 16667), 0, 50000) / 100000d;
        double r2 = ss * clamp(adj.value("adj2", 0), 0, 50000) / 100000d;
        Path2D path = new Path2D.Double();
        path.moveTo(x + r1, y);
        path.lineTo(x + w - r2, y);
        if (r2 > 0.01) {
            arcTo(path, r2, r2, 270, 90);
        } else {
            path.lineTo(x + w, y);
        }
        path.lineTo(x + w, y + h - r1);
        if (r1 > 0.01) {
            arcTo(path, r1, r1, 0, 90);
        } else {
            path.lineTo(x + w, y + h);
        }
        path.lineTo(x + r2, y + h);
        if (r2 > 0.01) {
            arcTo(path, r2, r2, 90, 90);
        } else {
            path.lineTo(x, y + h);
        }
        path.lineTo(x, y + r1);
        if (r1 > 0.01) {
            arcTo(path, r1, r1, 180, 90);
        } else {
            path.lineTo(x, y);
        }
        path.closePath();
        return path;
    }

    public static Shape snip1Rect(double x, double y, double w, double h, ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double r = ss * clamp(adj.value("adj", 16667), 0, 50000) / 100000d;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y);
        path.lineTo(x + w - r, y);
        path.lineTo(x + w, y + r);
        path.lineTo(x + w, y + h);
        path.lineTo(x, y + h);
        path.closePath();
        return path;
    }

    public static Shape snip2SameRect(double x, double y, double w, double h, ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double r1 = ss * clamp(adj.value("adj1", 16667), 0, 50000) / 100000d;
        double r2 = ss * clamp(adj.value("adj2", 0), 0, 50000) / 100000d;
        Path2D path = new Path2D.Double();
        path.moveTo(x + r1, y);
        path.lineTo(x + w - r1, y);
        path.lineTo(x + w, y + r1);
        path.lineTo(x + w, y + h - r2);
        path.lineTo(x + w - r2, y + h);
        path.lineTo(x + r2, y + h);
        path.lineTo(x, y + h - r2);
        path.lineTo(x, y + r1);
        path.closePath();
        return path;
    }

    public static Shape snip2DiagRect(double x, double y, double w, double h, ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double l = ss * clamp(adj.value("adj1", 0), 0, 50000) / 100000d;
        double r = ss * clamp(adj.value("adj2", 16667), 0, 50000) / 100000d;
        Path2D path = new Path2D.Double();
        path.moveTo(x + l, y);
        path.lineTo(x + w - r, y);
        path.lineTo(x + w, y + r);
        path.lineTo(x + w, y + h - l);
        path.lineTo(x + w - l, y + h);
        path.lineTo(x + r, y + h);
        path.lineTo(x, y + h - r);
        path.lineTo(x, y + l);
        path.closePath();
        return path;
    }

    public static Shape snipRoundRect(double x, double y, double w, double h, ShapeAdjustments adj) {
        double ss = Math.max(0.001, Math.min(w, h));
        double r1 = ss * clamp(adj.value("adj1", 16667), 0, 50000) / 100000d;
        double r2 = ss * clamp(adj.value("adj2", 16667), 0, 50000) / 100000d;
        Path2D path = new Path2D.Double();
        path.moveTo(x + r1, y);
        path.lineTo(x + w - r2, y);
        path.lineTo(x + w, y + r2);
        path.lineTo(x + w, y + h);
        path.lineTo(x, y + h);
        path.lineTo(x, y + r1);
        if (r1 > 0.01) {
            arcTo(path, r1, r1, 180, 90);
        } else {
            path.lineTo(x, y);
        }
        path.closePath();
        return path;
    }

    public static Shape triangle(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x, y + h);
        path.lineTo(x + w / 2d, y);
        path.lineTo(x + w, y + h);
        path.closePath();
        return path;
    }

    public static Shape rightTriangle(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x, y);
        path.lineTo(x + w, y + h);
        path.lineTo(x, y + h);
        path.closePath();
        return path;
    }

    public static Shape parallelogram(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x + w * 0.2, y);
        path.lineTo(x + w, y);
        path.lineTo(x + w * 0.8, y + h);
        path.lineTo(x, y + h);
        path.closePath();
        return path;
    }

    public static Shape trapezoid(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x + w * 0.2, y);
        path.lineTo(x + w * 0.8, y);
        path.lineTo(x + w, y + h);
        path.lineTo(x, y + h);
        path.closePath();
        return path;
    }

    public static Shape nonIsoscelesTrapezoid(double x, double y, double w, double h,
                                        ShapeAdjustments adj) {
        double ss = Math.min(w, h);
        double maxAdj = 50000d * w / ss;
        double a1 = clamp(adj.value("adj1", 25000), 0, maxAdj);
        double a2 = clamp(adj.value("adj2", 25000), 0, maxAdj);
        double x2 = ss * a1 / 100000d;
        double x3 = w - ss * a2 / 100000d;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y + h);
        path.lineTo(x + x2, y);
        path.lineTo(x + x3, y);
        path.lineTo(x + w, y + h);
        path.closePath();
        return path;
    }

    public static Shape regularPolygon(double x, double y, double w, double h, int sides) {
        Path2D path = new Path2D.Double();
        for (int i = 0; i < sides; i++) {
            double angle = -Math.PI / 2 + 2 * Math.PI * i / sides;
            double px = x + w / 2d + Math.cos(angle) * w / 2d;
            double py = y + h / 2d + Math.sin(angle) * h / 2d;
            if (i == 0) {
                path.moveTo(px, py);
            } else {
                path.lineTo(px, py);
            }
        }
        path.closePath();
        return path;
    }

    public static Shape octagon(double x, double y, double w, double h, ShapeAdjustments adj) {
        double ss = Math.min(w, h);
        double a = clamp(adj.value("adj", 29289), 0, 50000);
        double x1 = ss * a / 100000d;
        double x2 = w - x1;
        double y2 = h - x1;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y + x1);
        path.lineTo(x + x1, y);
        path.lineTo(x + x2, y);
        path.lineTo(x + w, y + x1);
        path.lineTo(x + w, y + y2);
        path.lineTo(x + x2, y + h);
        path.lineTo(x + x1, y + h);
        path.lineTo(x, y + y2);
        path.closePath();
        return path;
    }

    public static Shape pie(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        double cx = x + w / 2d;
        double cy = y + h / 2d;
        path.moveTo(cx + w / 2d, cy);
        arcTo(path, w / 2d, h / 2d, 0, 270);
        path.lineTo(cx, cy);
        path.closePath();
        return path;
    }

    public static Shape pieWedge(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x, y + h);
        arcTo(path, w, h, 180, 90);
        path.lineTo(x + w, y + h);
        path.closePath();
        return path;
    }

    public static Shape donut(double x, double y, double w, double h) {
        return donutArea(x, y, w, h);
    }

    public static Area donutArea(double x, double y, double w, double h) {
        double t = Math.min(w, h) * 0.18;
        Area area = new Area(new Ellipse2D.Double(x, y, w, h));
        area.subtract(new Area(new Ellipse2D.Double(x + t, y + t,
                Math.max(0.1, w - 2 * t), Math.max(0.1, h - 2 * t))));
        return area;
    }

    public static Shape noSmoking(double x, double y, double w, double h, ShapeAdjustments adj) {
        double ss = Math.min(w, h);
        double a = clamp(adj.value("adj", 18750), 0, 50000);
        double dr = ss * a / 100000d;
        double cx = x + w / 2d;
        double cy = y + h / 2d;
        Ellipse2D outer = new Ellipse2D.Double(x, y, w, h);
        Ellipse2D inner = new Ellipse2D.Double(
                x + dr, y + dr,
                Math.max(0.1, w - 2 * dr), Math.max(0.1, h - 2 * dr));
        Area area = new Area(outer);
        area.subtract(new Area(inner));

        double barWidth = Math.max(1, Math.max(dr * 1.4, ss * 0.28));
        double length = Math.hypot(w, h) * 2d;
        Shape barRect = AffineTransform.getRotateInstance(
                        Math.PI / 4d, cx, cy)
                .createTransformedShape(new Rectangle2D.Double(
                        cx - length / 2d, cy - barWidth / 2d,
                        length, barWidth));
        Area bar = new Area(barRect);
        bar.intersect(new Area(outer));
        area.add(bar);
        return area;
    }

    public static Shape blockArc(double x, double y, double w, double h, ShapeAdjustments adj) {
        double stAng = clamp(adj.value("adj1", 10800000),
                0, 21599999) / 60000d;
        double istAng = clamp(adj.value("adj2", 0),
                0, 21599999) / 60000d;
        double a3 = clamp(adj.value("adj3", 25000), 0, 50000);
        double sw1 = istAng - stAng;
        double swAng = sw1 > 0 ? sw1 : sw1 + 360d;
        double dr = Math.min(w, h) * a3 / 100000d;
        double iwd2 = w / 2d - dr;
        double ihd2 = h / 2d - dr;
        double cx = x + w / 2d;
        double cy = y + h / 2d;
        double st = Math.toRadians(stAng);
        double ist = Math.toRadians(istAng);

        Path2D path = new Path2D.Double();
        path.moveTo(cx + w / 2d * Math.cos(st), cy + h / 2d * Math.sin(st));
        arcTo(path, w / 2d, h / 2d, stAng, swAng);
        path.lineTo(cx + iwd2 * Math.cos(ist), cy + ihd2 * Math.sin(ist));
        arcTo(path, iwd2, ihd2, istAng, -swAng);
        path.closePath();
        return path;
    }

    public static Shape chord(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        double cx = x + w / 2d;
        double cy = y + h / 2d;
        double rx = w / 2d;
        double ry = h / 2d;
        double start = Math.toRadians(45);
        path.moveTo(cx + rx * Math.cos(start), cy + ry * Math.sin(start));
        arcTo(path, rx, ry, 45, 225);
        path.closePath();
        return path;
    }

    public static Shape chevron(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x, y);
        path.lineTo(x + w * 0.75, y);
        path.lineTo(x + w, y + h / 2d);
        path.lineTo(x + w * 0.75, y + h);
        path.lineTo(x, y + h);
        path.lineTo(x + w * 0.25, y + h / 2d);
        path.closePath();
        return path;
    }

    public static Shape plaque(double x, double y, double w, double h, ShapeAdjustments adj) {
        double ss = Math.min(w, h);
        double a = clamp(adj.value("adj", 16667), 0, 50000);
        double x1 = ss * a / 100000d;
        double x2 = w - x1;
        double y2 = h - x1;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y + x1);
        arcTo(path, x1, x1, 90, -90);
        path.lineTo(x + x2, y);
        arcTo(path, x1, x1, 180, -90);
        path.lineTo(x + w, y + y2);
        arcTo(path, x1, x1, 270, -90);
        path.lineTo(x + x1, y + h);
        arcTo(path, x1, x1, 0, -90);
        path.closePath();
        return path;
    }

    public static Shape teardrop(double x, double y, double w, double h, ShapeAdjustments adj) {
        double hc = x + w / 2d;
        double vc = y + h / 2d;
        double a = clamp(adj.value("adj", 100000), 0, 200000);
        double r2 = Math.sqrt(2d);
        double tw = w / 2d * r2;
        double th = h / 2d * r2;
        double sw = tw * a / 100000d;
        double sh = th * a / 100000d;
        double dx1 = sw * Math.cos(Math.PI / 4d);
        double dy1 = sh * Math.sin(Math.PI / 4d);
        double x1 = hc + dx1;
        double y1 = vc - dy1;
        double x2 = (hc + x1) / 2d;
        double y2 = (vc + y1) / 2d;

        Path2D path = new Path2D.Double();
        path.moveTo(x, vc);
        arcTo(path, w / 2d, h / 2d, 180, 90);
        path.quadTo(x2, y, x1, y1);
        path.quadTo(x + w, y2, x + w, vc);
        arcTo(path, w / 2d, h / 2d, 0, 90);
        arcTo(path, w / 2d, h / 2d, 90, 90);
        path.closePath();
        return path;
    }

    public static Shape heart(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x + w / 2d, y + h);
        path.curveTo(x - w * 0.05, y + h * 0.6,
                x + w * 0.12, y, x + w * 0.5, y + h * 0.28);
        path.curveTo(x + w * 0.88, y,
                x + w * 1.05, y + h * 0.6, x + w / 2d, y + h);
        path.closePath();
        return path;
    }

    public static Shape cloud(double x, double y, double w, double h) {
        Area area = new Area(new Ellipse2D.Double(
                x + w * 0.12, y + h * 0.35, w * 0.45, h * 0.55));
        area.add(new Area(new Ellipse2D.Double(
                x + w * 0.3, y + h * 0.15, w * 0.5, h * 0.65)));
        area.add(new Area(new Ellipse2D.Double(
                x + w * 0.6, y + h * 0.3, w * 0.35, h * 0.5)));
        area.add(new Area(new RoundRectangle2D.Double(
                x + w * 0.1, y + h * 0.55, w * 0.8, h * 0.35,
                w * 0.2, h * 0.2)));
        return area;
    }

    public static Shape wave(double x, double y, double w, double h, ShapeAdjustments adj) {
        double a1 = clamp(adj.value("adj1", 12500), 0, 20000);
        double a2 = clamp(adj.value("adj2", 0), -10000, 10000);
        double y1 = h * a1 / 100000d;
        double dy2 = y1 * 10d / 3d;
        double y2 = y1 - dy2;
        double y3 = y1 + dy2;
        double y4 = h - y1;
        double y5 = y4 - dy2;
        double y6 = y4 + dy2;
        double of2 = w * a2 / 50000d;
        double dx2 = of2 > 0 ? 0 : of2;
        double x2 = -dx2;
        double dx5 = of2 > 0 ? of2 : 0;
        double x5 = w - dx5;
        double dx3 = (dx2 + x5) / 3d;
        double x3 = x2 + dx3;
        double x4 = (x3 + x5) / 2d;
        double x6 = dx5;
        double x10 = w + dx2;
        double x7 = x6 + dx3;
        double x8 = (x7 + x10) / 2d;

        Path2D path = new Path2D.Double();
        path.moveTo(x + x2, y + y1);
        path.curveTo(x + x3, y + y2, x + x4, y + y3, x + x5, y + y1);
        path.lineTo(x + x10, y + y4);
        path.curveTo(x + x8, y + y6, x + x7, y + y5, x + x6, y + y4);
        path.closePath();
        return path;
    }

    public static Shape doubleWave(double x, double y, double w, double h, ShapeAdjustments adj) {
        double a1 = clamp(adj.value("adj1", 6250), 0, 12500);
        double a2 = clamp(adj.value("adj2", 0), -10000, 10000);
        double y1 = h * a1 / 100000d;
        double dy2 = y1 * 10d / 3d;
        double y2 = y1 - dy2;
        double y3 = y1 + dy2;
        double y4 = h - y1;
        double y5 = y4 - dy2;
        double y6 = y4 + dy2;
        double dx1 = w * a2 / 100000d;
        double of2 = w * a2 / 50000d;
        double dx2 = of2 > 0 ? 0 : of2;
        double x2 = -dx2;
        double dx8 = of2 > 0 ? of2 : 0;
        double x8 = w - dx8;
        double dx3 = (dx2 + x8) / 6d;
        double x3 = x2 + dx3;
        double dx4 = (dx2 + x8) / 3d;
        double x4 = x2 + dx4;
        double x5 = (x2 + x8) / 2d;
        double x6 = x5 + dx3;
        double x7 = (x6 + x8) / 2d;
        double x9 = dx8;
        double x15 = w + dx2;
        double x10 = x9 + dx3;
        double x11 = x9 + dx4;
        double x12 = (x9 + x15) / 2d;
        double x13 = x12 + dx3;
        double x14 = (x13 + x15) / 2d;

        Path2D path = new Path2D.Double();
        path.moveTo(x + x2, y + y1);
        path.curveTo(x + x3, y + y2, x + x4, y + y3, x + x5, y + y1);
        path.curveTo(x + x6, y + y2, x + x7, y + y3, x + x8, y + y1);
        path.lineTo(x + x15, y + y4);
        path.curveTo(x + x14, y + y6, x + x13, y + y5, x + x12, y + y4);
        path.curveTo(x + x11, y + y6, x + x10, y + y5, x + x9, y + y4);
        path.closePath();
        return path;
    }

    public static Shape plus(double x, double y, double w, double h, ShapeAdjustments adj) {
        double ss = Math.min(w, h);
        double a = clamp(adj.value("adj", 25000), 0, 50000);
        double x1 = ss * a / 100000d;
        double x2 = w - x1;
        double y2 = h - x1;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y + x1);
        path.lineTo(x + x1, y + x1);
        path.lineTo(x + x1, y);
        path.lineTo(x + x2, y);
        path.lineTo(x + x2, y + x1);
        path.lineTo(x + w, y + x1);
        path.lineTo(x + w, y + y2);
        path.lineTo(x + x2, y + y2);
        path.lineTo(x + x2, y + h);
        path.lineTo(x + x1, y + h);
        path.lineTo(x + x1, y + y2);
        path.lineTo(x, y + y2);
        path.closePath();
        return path;
    }

    public static Shape mathPlus(double x, double y, double w, double h) {
        Area area = new Area(new Rectangle2D.Double(
                x + w * 0.4, y, w * 0.2, h));
        area.add(new Area(new Rectangle2D.Double(
                x, y + h * 0.4, w, h * 0.2)));
        return area;
    }

    public static Shape mathMinus(double x, double y, double w, double h) {
        return new Rectangle2D.Double(x, y + h * 0.4, w, h * 0.2);
    }

    public static Shape mathMultiply(double x, double y, double w, double h, ShapeAdjustments adj) {
        if (w < 1 || h < 1) {
            return new java.awt.geom.Line2D.Double(x, y, x + w, y + h);
        }
        double a1 = clamp(adj.value("adj1", 23520), 0, 51965);
        double th = Math.min(w, h) * a1 / 100000d;
        double ang = Math.atan2(h, w);
        double sa = Math.sin(ang);
        double ca = Math.cos(ang);
        double ta = Math.tan(ang);
        double dl = Math.hypot(w, h);
        double rw = dl * 51965d / 100000d;
        double lM = dl - rw;
        double xM = ca * lM / 2d;
        double yM = sa * lM / 2d;
        double dxAM = sa * th / 2d;
        double dyAM = ca * th / 2d;
        double xA = xM - dxAM;
        double yA = yM + dyAM;
        double xB = xM + dxAM;
        double yB = yM - dyAM;
        double xBC = w / 2d - xB;
        double yBC = xBC * ta;
        double yC = yBC + yB;
        double xD = w - xB;
        double xE = w - xA;
        double yFE = h / 2d - yA;
        double xFE = yFE / ta;
        double xF = xE - xFE;
        double xL = xA + xFE;
        double yG = h - yA;
        double yH = h - yB;
        double yI = h - yC;

        Path2D path = new Path2D.Double();
        path.moveTo(x + xA, y + yA);
        path.lineTo(x + xB, y + yB);
        path.lineTo(x + w / 2d, y + yC);
        path.lineTo(x + xD, y + yB);
        path.lineTo(x + xE, y + yA);
        path.lineTo(x + xF, y + h / 2d);
        path.lineTo(x + xE, y + yG);
        path.lineTo(x + xD, y + yH);
        path.lineTo(x + w / 2d, y + yI);
        path.lineTo(x + xB, y + yH);
        path.lineTo(x + xA, y + yG);
        path.lineTo(x + xL, y + h / 2d);
        path.closePath();
        return path;
    }

    public static Shape mathDivide(double x, double y, double w, double h) {
        Area area = new Area(mathMinus(x, y, w, h));
        double r = Math.min(w, h) * 0.1;
        area.add(new Area(new Ellipse2D.Double(
                x + w / 2d - r, y + h * 0.18 - r, 2 * r, 2 * r)));
        area.add(new Area(new Ellipse2D.Double(
                x + w / 2d - r, y + h * 0.82 - r, 2 * r, 2 * r)));
        return area;
    }

    public static Shape mathEqual(double x, double y, double w, double h) {
        Area area = new Area(new Rectangle2D.Double(
                x, y + h * 0.3, w, h * 0.15));
        area.add(new Area(new Rectangle2D.Double(
                x, y + h * 0.55, w, h * 0.15)));
        return area;
    }

    public static Shape mathNotEqual(double x, double y, double w, double h) {
        Area area = new Area(mathEqual(x, y, w, h));
        area.add(new Area(strokedShape(new java.awt.geom.Line2D.Double(
                x + w * 0.2, y + h * 0.75,
                x + w * 0.8, y + h * 0.25), Math.min(w, h) * 0.08)));
        return area;
    }

    public static Shape diagonalStripe(double x, double y, double w, double h, ShapeAdjustments adj) {
        double a = clamp(adj.value("adj", 50000), 0, 100000);
        double x2 = w * a / 100000d;
        double y2 = h * a / 100000d;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y + y2);
        path.lineTo(x + x2, y);
        path.lineTo(x + w, y);
        path.lineTo(x, y + h);
        path.closePath();
        return path;
    }

    public static Shape corner(double x, double y, double w, double h) {
        Area area = new Area(new Rectangle2D.Double(
                x, y, w * 0.3, h));
        area.add(new Area(new Rectangle2D.Double(
                x, y + h * 0.7, w, h * 0.3)));
        return area;
    }

    public static Shape cornerTabs(double x, double y, double w, double h) {
        double dx = Math.hypot(w, h) / 20d;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y);
        path.lineTo(x + dx, y);
        path.lineTo(x, y + dx);
        path.closePath();
        path.moveTo(x, y + h - dx);
        path.lineTo(x + dx, y + h);
        path.lineTo(x, y + h);
        path.closePath();
        path.moveTo(x + w - dx, y);
        path.lineTo(x + w, y);
        path.lineTo(x + w, y + dx);
        path.closePath();
        path.moveTo(x + w, y + h - dx);
        path.lineTo(x + w, y + h);
        path.lineTo(x + w - dx, y + h);
        path.closePath();
        return path;
    }

    public static Shape squareTabs(double x, double y, double w, double h) {
        double dx = Math.hypot(w, h) / 20d;
        Path2D path = new Path2D.Double();
        path.append(new Rectangle2D.Double(x, y, dx, dx), false);
        path.append(new Rectangle2D.Double(x, y + h - dx, dx, dx), false);
        path.append(new Rectangle2D.Double(x + w - dx, y, dx, dx), false);
        path.append(new Rectangle2D.Double(x + w - dx, y + h - dx, dx, dx), false);
        return path;
    }

    public static Shape plaqueTabs(double x, double y, double w, double h) {
        double dx = Math.hypot(w, h) / 20d;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y);
        path.lineTo(x + dx, y);
        arcTo(path, dx, dx, 0, 90);
        path.closePath();
        path.moveTo(x, y + h - dx);
        arcTo(path, dx, dx, 270, 90);
        path.lineTo(x, y + h);
        path.closePath();
        path.moveTo(x + w, y);
        path.lineTo(x + w, y + dx);
        arcTo(path, dx, dx, 90, 90);
        path.closePath();
        path.moveTo(x + w - dx, y + h);
        arcTo(path, dx, dx, 180, 90);
        path.lineTo(x + w, y + h);
        path.closePath();
        return path;
    }

    public static Shape frame(double x, double y, double w, double h) {
        Area area = new Area(new Rectangle2D.Double(x, y, w, h));
        area.subtract(new Area(new Rectangle2D.Double(
                x + w * 0.18, y + h * 0.18, w * 0.64, h * 0.64)));
        return area;
    }

    public static Shape halfFrame(double x, double y, double w, double h, ShapeAdjustments adj) {
        double ss = Math.min(w, h);
        double maxAdj2 = 100000d * w / ss;
        double a2 = clamp(adj.value("adj2", 33333), 0, maxAdj2);
        double x1 = ss * a2 / 100000d;
        double g1 = h * x1 / w;
        double g2 = h - g1;
        double maxAdj1 = 100000d * g2 / ss;
        double a1 = clamp(adj.value("adj1", 33333), 0, maxAdj1);
        double y1 = ss * a1 / 100000d;
        double dx2 = y1 * w / h;
        double x2 = w - dx2;
        double dy2 = x1 * h / w;
        double y2 = h - dy2;

        Path2D path = new Path2D.Double();
        path.moveTo(x, y);
        path.lineTo(x + w, y);
        path.lineTo(x + x2, y + y1);
        path.lineTo(x + x1, y + y1);
        path.lineTo(x + x1, y + y2);
        path.lineTo(x, y + h);
        path.closePath();
        return path;
    }

    public static Shape cube(double x, double y, double w, double h, ShapeAdjustments adj) {
        double ss = Math.min(w, h);
        double a = clamp(adj.value("adj", 25000), 0, 100000);
        double y1 = ss * a / 100000d;
        double y4 = h - y1;
        double x4 = w - y1;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y + y1);
        path.lineTo(x + y1, y);
        path.lineTo(x + w, y);
        path.lineTo(x + w, y + y4);
        path.lineTo(x + x4, y + h);
        path.lineTo(x, y + h);
        path.closePath();
        return path;
    }

    public static Shape can(double x, double y, double w, double h, ShapeAdjustments adj) {
        double ss = Math.min(w, h);
        double maxAdj = 50000d * h / ss;
        double a = clamp(adj.value("adj", 25000), 0, maxAdj);
        double y1 = ss * a / 200000d;
        double y3 = h - y1;
        Path2D path = new Path2D.Double();
        path.moveTo(x + w, y + y1);
        arcTo(path, w / 2d, y1, 0, 180);
        arcTo(path, w / 2d, y1, 180, 180);
        path.lineTo(x + w, y + y3);
        arcTo(path, w / 2d, y1, 0, 180);
        path.lineTo(x, y + y1);
        return path;
    }

    public static Shape funnel(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.moveTo(x, y);
        path.lineTo(x + w, y);
        path.lineTo(x + w * 0.65, y + h * 0.45);
        path.lineTo(x + w * 0.65, y + h);
        path.lineTo(x + w * 0.35, y + h);
        path.lineTo(x + w * 0.35, y + h * 0.45);
        path.closePath();
        return path;
    }

    public static Shape gear(double x, double y, double w, double h, int teeth) {
        double cx = x + w / 2d;
        double cy = y + h / 2d;
        double outer = Math.min(w, h) * 0.48;
        double inner = outer * 0.72;
        Path2D path = new Path2D.Double();
        int points = teeth * 4;
        for (int i = 0; i < points; i++) {
            double angle = -Math.PI / 2 + 2 * Math.PI * i / points;
            double radius = (i % 4 == 0 || i % 4 == 1) ? outer : inner;
            double px = cx + Math.cos(angle) * radius * w / Math.min(w, h);
            double py = cy + Math.sin(angle) * radius * h / Math.min(w, h);
            if (i == 0) {
                path.moveTo(px, py);
            } else {
                path.lineTo(px, py);
            }
        }
        path.closePath();
        Area area = new Area(path);
        double hole = Math.min(w, h) * 0.12;
        area.subtract(new Area(new Ellipse2D.Double(
                cx - hole, cy - hole, hole * 2, hole * 2)));
        return area;
    }

    public static Shape lightningBolt(double x, double y, double w, double h) {
        double[] xs = {8472, 12860, 11050, 16577, 14767, 21600,
                10012, 12222, 5022, 7602, 0};
        double[] ys = {0, 6080, 6797, 12007, 12877, 21600,
                14915, 13987, 9705, 8382, 3890};
        Path2D path = new Path2D.Double();
        for (int i = 0; i < xs.length; i++) {
            double px = x + xs[i] * w / 21600d;
            double py = y + ys[i] * h / 21600d;
            if (i == 0) {
                path.moveTo(px, py);
            } else {
                path.lineTo(px, py);
            }
        }
        path.closePath();
        return path;
    }

    public static Shape moon(double x, double y, double w, double h) {
        Area area = new Area(new Ellipse2D.Double(x, y, w, h));
        area.subtract(new Area(new Ellipse2D.Double(
                x + w * 0.28, y - h * 0.05, w * 0.9, h * 0.9)));
        return area;
    }

    public static Shape sun(double x, double y, double w, double h) {
        Area area = new Area(new Ellipse2D.Double(
                x + w * 0.2, y + h * 0.2, w * 0.6, h * 0.6));
        double cx = x + w / 2d;
        double cy = y + h / 2d;
        for (int i = 0; i < 8; i++) {
            double a = 2 * Math.PI * i / 8;
            double px = cx + Math.cos(a) * w * 0.45;
            double py = cy + Math.sin(a) * h * 0.45;
            double bx = cx + Math.cos(a - 0.12) * w * 0.28;
            double by = cy + Math.sin(a - 0.12) * h * 0.28;
            double ex = cx + Math.cos(a + 0.12) * w * 0.28;
            double ey = cy + Math.sin(a + 0.12) * h * 0.28;
            Path2D ray = new Path2D.Double();
            ray.moveTo(bx, by);
            ray.lineTo(px, py);
            ray.lineTo(ex, ey);
            ray.closePath();
            area.add(new Area(ray));
        }
        return area;
    }

    public static Shape arcShape(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        double cx = x + w / 2d;
        path.moveTo(cx, y);
        arcTo(path, w / 2d, h / 2d, 270, 90);
        return path;
    }

    public static Shape bracket(double x, double y, double w, double h, boolean left) {
        Path2D path = new Path2D.Double();
        if (left) {
            path.moveTo(x + w * 0.75, y);
            path.lineTo(x + w * 0.25, y);
            path.lineTo(x + w * 0.25, y + h);
            path.lineTo(x + w * 0.75, y + h);
        } else {
            path.moveTo(x + w * 0.25, y);
            path.lineTo(x + w * 0.75, y);
            path.lineTo(x + w * 0.75, y + h);
            path.lineTo(x + w * 0.25, y + h);
        }
        return path;
    }

    public static Shape bracketPair(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.append(bracket(x, y, w / 2d, h, true), false);
        path.append(bracket(x + w / 2d, y, w / 2d, h, false), false);
        return path;
    }

    public static Shape brace(double x, double y, double w, double h, boolean left) {
        Path2D path = new Path2D.Double();
        if (left) {
            path.moveTo(x + w * 0.8, y);
            path.curveTo(x + w * 0.55, y,
                    x + w * 0.55, y + h * 0.2, x + w * 0.5, y + h * 0.4);
            path.curveTo(x + w * 0.45, y + h * 0.46,
                    x + w * 0.2, y + h * 0.46, x + w * 0.2, y + h * 0.5);
            path.curveTo(x + w * 0.2, y + h * 0.54,
                    x + w * 0.45, y + h * 0.54, x + w * 0.5, y + h * 0.6);
            path.curveTo(x + w * 0.55, y + h * 0.8,
                    x + w * 0.55, y + h, x + w * 0.8, y + h);
        } else {
            path.moveTo(x + w * 0.2, y);
            path.curveTo(x + w * 0.45, y,
                    x + w * 0.45, y + h * 0.2, x + w * 0.5, y + h * 0.4);
            path.curveTo(x + w * 0.55, y + h * 0.46,
                    x + w * 0.8, y + h * 0.46, x + w * 0.8, y + h * 0.5);
            path.curveTo(x + w * 0.8, y + h * 0.54,
                    x + w * 0.55, y + h * 0.54, x + w * 0.5, y + h * 0.6);
            path.curveTo(x + w * 0.45, y + h * 0.8,
                    x + w * 0.45, y + h, x + w * 0.2, y + h);
        }
        return path;
    }

    public static Shape bracePair(double x, double y, double w, double h) {
        Path2D path = new Path2D.Double();
        path.append(brace(x, y, w / 2d, h, true), false);
        path.append(brace(x + w / 2d, y, w / 2d, h, false), false);
        return path;
    }

    public static Shape ribbon(double x, double y, double w, double h, boolean inverted) {
        Path2D band = new Path2D.Double();
        band.moveTo(x + w * 0.12, y + h * 0.20);
        band.lineTo(x + w * 0.88, y + h * 0.20);
        band.lineTo(x + w * 0.88, y + h * 0.72);
        band.lineTo(x + w * 0.12, y + h * 0.72);
        band.closePath();
        Area area = new Area(band);

        Path2D left = new Path2D.Double();
        left.moveTo(x, y + h * 0.32);
        left.lineTo(x + w * 0.14, y + h * 0.38);
        left.lineTo(x + w * 0.14, y + h * 0.88);
        left.lineTo(x + w * 0.07, y + h * 0.78);
        left.lineTo(x, y + h * 0.88);
        left.closePath();
        area.add(new Area(left));

        Path2D right = new Path2D.Double();
        right.moveTo(x + w, y + h * 0.32);
        right.lineTo(x + w * 0.86, y + h * 0.38);
        right.lineTo(x + w * 0.86, y + h * 0.88);
        right.lineTo(x + w * 0.93, y + h * 0.78);
        right.lineTo(x + w, y + h * 0.88);
        right.closePath();
        area.add(new Area(right));

        return inverted ? flipVertical(area, y + h / 2d) : area;
    }

    public static Shape ellipseRibbon(double x, double y, double w, double h,
                                boolean inverted) {
        Path2D band = new Path2D.Double();
        band.moveTo(x, y + h * 0.55);
        band.quadTo(x + w / 2d, y - h * 0.02, x + w, y + h * 0.55);
        band.lineTo(x + w, y + h * 0.78);
        band.quadTo(x + w / 2d, y + h * 0.21, x, y + h * 0.78);
        band.closePath();
        Area area = new Area(band);

        Path2D left = new Path2D.Double();
        left.moveTo(x, y + h * 0.55);
        left.lineTo(x + w * 0.08, y + h * 0.63);
        left.lineTo(x + w * 0.08, y + h * 0.90);
        left.lineTo(x, y + h * 0.78);
        left.closePath();
        area.add(new Area(left));

        Path2D right = new Path2D.Double();
        right.moveTo(x + w, y + h * 0.55);
        right.lineTo(x + w * 0.92, y + h * 0.63);
        right.lineTo(x + w * 0.92, y + h * 0.90);
        right.lineTo(x + w, y + h * 0.78);
        right.closePath();
        area.add(new Area(right));

        return inverted ? flipVertical(area, y + h / 2d) : area;
    }

    public static Shape leftRightRibbon(double x, double y, double w, double h) {
        Path2D band = new Path2D.Double();
        band.moveTo(x + w * 0.18, y + h * 0.25);
        band.lineTo(x + w * 0.82, y + h * 0.25);
        band.lineTo(x + w * 0.82, y + h * 0.75);
        band.lineTo(x + w * 0.18, y + h * 0.75);
        band.closePath();
        Area area = new Area(band);

        Path2D left = new Path2D.Double();
        left.moveTo(x, y + h * 0.35);
        left.lineTo(x + w * 0.18, y + h * 0.25);
        left.lineTo(x + w * 0.18, y + h * 0.75);
        left.lineTo(x, y + h * 0.65);
        left.lineTo(x + w * 0.08, y + h * 0.50);
        left.closePath();
        area.add(new Area(left));

        Path2D right = new Path2D.Double();
        right.moveTo(x + w, y + h * 0.35);
        right.lineTo(x + w * 0.82, y + h * 0.25);
        right.lineTo(x + w * 0.82, y + h * 0.75);
        right.lineTo(x + w, y + h * 0.65);
        right.lineTo(x + w * 0.92, y + h * 0.50);
        right.closePath();
        area.add(new Area(right));
        return area;
    }

    public static Shape scroll(double x, double y, double w, double h, boolean vertical) {
        Area area = new Area();
        if (vertical) {
            area.add(new Area(new RoundRectangle2D.Double(
                    x + w * 0.18, y + h * 0.08, w * 0.64, h * 0.84,
                    w * 0.18, h * 0.18)));
            area.add(new Area(new RoundRectangle2D.Double(
                    x + w * 0.08, y + h * 0.02, w * 0.84, h * 0.18,
                    w * 0.16, h * 0.16)));
            area.add(new Area(new RoundRectangle2D.Double(
                    x + w * 0.08, y + h * 0.80, w * 0.84, h * 0.18,
                    w * 0.16, h * 0.16)));
        } else {
            area.add(new Area(new RoundRectangle2D.Double(
                    x + w * 0.08, y + h * 0.18, w * 0.84, h * 0.64,
                    w * 0.18, h * 0.18)));
            area.add(new Area(new RoundRectangle2D.Double(
                    x + w * 0.02, y + h * 0.08, w * 0.18, h * 0.84,
                    w * 0.16, h * 0.16)));
            area.add(new Area(new RoundRectangle2D.Double(
                    x + w * 0.80, y + h * 0.08, w * 0.18, h * 0.84,
                    w * 0.16, h * 0.16)));
        }
        return area;
    }

    public static Shape foldedCorner(double x, double y, double w, double h, double adjust) {
        double a = clamp(adjust, 0, 50000);
        double dy2 = Math.min(w, h) * a / 100000.0;
        double x1 = x + w - dy2;
        double y2 = y + h - dy2;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y);
        path.lineTo(x + w, y);
        path.lineTo(x + w, y2);
        path.lineTo(x1, y + h);
        path.lineTo(x, y + h);
        path.closePath();
        return path;
    }

    public static Shape irregularSeal1(double x, double y, double w, double h) {
        double[] radii = {1.0, 0.55, 0.85, 0.5, 0.95, 0.6, 1.0, 0.5,
                0.8, 0.55, 0.9, 0.45, 0.98, 0.6, 0.82, 0.5};
        double cx = x + w / 2d;
        double cy = y + h / 2d;
        Path2D path = new Path2D.Double();
        for (int i = 0; i < radii.length; i++) {
            double angle = -Math.PI / 2 + 2 * Math.PI * i / radii.length;
            double px = cx + Math.cos(angle) * w / 2d * radii[i];
            double py = cy + Math.sin(angle) * h / 2d * radii[i];
            if (i == 0) {
                path.moveTo(px, py);
            } else {
                path.lineTo(px, py);
            }
        }
        path.closePath();
        return path;
    }

    public static Shape star(double x, double y, double w, double h, int points) {
        Path2D path = new Path2D.Double();        double cx = x + w / 2;
        double cy = y + h / 2;
        double outer = Math.min(w, h) / 2;
        double inner = outer * 0.45;
        for (int i = 0; i < points * 2; i++) {
            double radius = i % 2 == 0 ? outer : inner;
            double angle = -Math.PI / 2 + i * Math.PI / points;
            double px = cx + Math.cos(angle) * radius * (w / Math.min(w, h));
            double py = cy + Math.sin(angle) * radius * (h / Math.min(w, h));
            if (i == 0) {
                path.moveTo(px, py);
            } else {
                path.lineTo(px, py);
            }
        }
        path.closePath();
        return path;
    }

    /**
     * ECMA-376 wedgeRectCallout silhouette. The tail is inserted into the
     * nearest side; the base of the wedge is not stroked.
     */
    public static Shape wedgeRectCallout(double x, double y, double w, double h,
                                   double adj1, double adj2) {
        if (w <= 0 || h <= 0) {
            return new Rectangle2D.Double(x, y, w, h);
        }
        double dxPos = w * adj1 / 100000d;
        double dyPos = h * adj2 / 100000d;
        double xPos = x + w / 2d + dxPos;
        double yPos = y + h / 2d + dyPos;
        double dq = dxPos * h / w;
        double dz = Math.abs(dyPos) - Math.abs(dq);
        double xg1 = dxPos > 0 ? 7 : 2;
        double xg2 = dxPos > 0 ? 10 : 5;
        double x1 = x + w * xg1 / 12d;
        double x2 = x + w * xg2 / 12d;
        double yg1 = dyPos > 0 ? 7 : 2;
        double yg2 = dyPos > 0 ? 10 : 5;
        double y1 = y + h * yg1 / 12d;
        double y2 = y + h * yg2 / 12d;
        double xl = dz > 0 ? x : (dxPos > 0 ? x : xPos);
        double xt = dz > 0 ? (dyPos > 0 ? x1 : xPos) : x1;
        double xr = dz > 0 ? x + w : (dxPos > 0 ? xPos : x + w);
        double xb = dz > 0 ? (dyPos > 0 ? xPos : x1) : x1;
        double yl = dz > 0 ? y1 : (dxPos > 0 ? y1 : yPos);
        double yt = dz > 0 ? (dyPos > 0 ? y : yPos) : y;
        double yr = dz > 0 ? y1 : (dxPos > 0 ? yPos : y1);
        double yb = dz > 0 ? (dyPos > 0 ? yPos : y + h) : y + h;
        Path2D path = new Path2D.Double();
        path.moveTo(x, y);
        path.lineTo(x1, y);
        path.lineTo(xt, yt);
        path.lineTo(x2, y);
        path.lineTo(x + w, y);
        path.lineTo(x + w, y1);
        path.lineTo(xr, yr);
        path.lineTo(x + w, y2);
        path.lineTo(x + w, y + h);
        path.lineTo(x2, y + h);
        path.lineTo(xb, yb);
        path.lineTo(x1, y + h);
        path.lineTo(x, y + h);
        path.lineTo(x, y2);
        path.lineTo(xl, yl);
        path.lineTo(x, y1);
        path.closePath();
        return path;
    }

    public static Shape arrowCallout(double x, double y, double w, double h,
                               ShapeAdjustments adj, int direction) {
        Area area = new Area();
        double a1 = adj.value("adj1", 50000);
        double a2 = adj.value("adj2", 50000);
        switch (direction) {
            case 1 -> {
                area.add(new Area(horizontalArrow(x, y, w, h, adj, true)));
                area.add(new Area(new Rectangle2D.Double(
                        x + w * 0.62, y + h * 0.2, w * 0.38, h * 0.6)));
            }
            case 2 -> {
                area.add(new Area(horizontalArrow(x, y, w, h, adj, false)));
                area.add(new Area(new Rectangle2D.Double(
                        x, y + h * 0.2, w * 0.38, h * 0.6)));
            }
            case 3 -> {
                area.add(new Area(verticalArrow(x, y, w, h, adj, true)));
                area.add(new Area(new Rectangle2D.Double(
                        x + w * 0.2, y + h * 0.62, w * 0.6, h * 0.38)));
            }
            case 4 -> {
                area.add(new Area(verticalArrow(x, y, w, h, adj, false)));
                area.add(new Area(new Rectangle2D.Double(
                        x + w * 0.2, y, w * 0.6, h * 0.38)));
            }
            case 5 -> {
                area.add(new Area(leftRightArrow(x, y, w, h, a1, a2)));
                area.add(new Area(new Rectangle2D.Double(
                        x + w * 0.30, y + h * 0.2, w * 0.4, h * 0.6)));
            }
            case 6 -> {
                area.add(new Area(upDownArrow(x, y, w, h, adj)));
                area.add(new Area(new Rectangle2D.Double(
                        x + w * 0.2, y + h * 0.30, w * 0.6, h * 0.4)));
            }
            default -> {
                area.add(new Area(quadArrow(x, y, w, h, adj)));
                area.add(new Area(new Rectangle2D.Double(
                        x + w * 0.3, y + h * 0.3, w * 0.4, h * 0.4)));
            }
        }
        return area;
    }

    public static Shape callout(double x, double y, double w, double h, double arc) {
        Path2D path = new Path2D.Double();
        path.append(new RoundRectangle2D.Double(x, y, w, h * 0.82, arc, arc), false);
        path.moveTo(x + w * 0.25, y + h * 0.82);
        path.lineTo(x + w * 0.12, y + h);
        path.lineTo(x + w * 0.42, y + h * 0.82);
        path.closePath();
        return path;
    }

    public static Shape strokedShape(Shape centerline, double thickness) {
        return new BasicStroke((float) Math.max(0.5, thickness),
                BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND)
                .createStrokedShape(centerline);
    }

    private static double clamp(double value, double min, double max) {
        return Math.max(min, Math.min(max, value));
    }
}
