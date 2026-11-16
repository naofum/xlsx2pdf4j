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

import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.xml.sax.InputSource;

import javax.xml.parsers.DocumentBuilderFactory;
import java.awt.Color;
import java.awt.LinearGradientPaint;
import java.awt.Paint;
import java.awt.RadialGradientPaint;
import java.awt.geom.Point2D;
import java.awt.geom.Rectangle2D;
import java.io.StringReader;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

/**
 * SpreadsheetML gradient cell fill. Excel stores cell gradients in styles.xml
 * as a gradientFill with stop/color children; the DrawingML classes cannot be
 * used from poi-ooxml-lite, so the fill fragment is read as XML.
 */
final class CellGradient {

    private static final String NS =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private final boolean path;
    private final double degree;
    private final double left;
    private final double right;
    private final double top;
    private final double bottom;
    private final List<Stop> stops;

    private CellGradient(boolean path, double degree, double left, double right,
                         double top, double bottom, List<Stop> stops) {
        this.path = path;
        this.degree = degree;
        this.left = left;
        this.right = right;
        this.top = top;
        this.bottom = bottom;
        this.stops = List.copyOf(stops);
    }

    record Stop(double position, Color color) {
    }

    static CellGradient read(XSSFWorkbook workbook, XSSFCellStyle style) {
        if (workbook == null || style == null || style.getCoreXf() == null) {
            return null;
        }
        try {
            int fillId = (int) style.getCoreXf().getFillId();
            XSSFCellFill fill = workbook.getStylesSource() == null
                    ? null : workbook.getStylesSource().getFillAt(fillId);
            if (fill == null || fill.getCTFill() == null) {
                return null;
            }
            String fragment = fill.getCTFill().xmlText();
            if (!fragment.contains("gradientFill")) {
                return null;
            }
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            Document document = factory.newDocumentBuilder().parse(new InputSource(
                    new StringReader("<fill xmlns=\"" + NS + "\">" + fragment + "</fill>")));
            Element gradient = find(document.getDocumentElement(), "gradientFill");
            if (gradient == null) {
                return null;
            }
            List<Stop> stops = new ArrayList<>();
            for (Node node = gradient.getFirstChild(); node != null; node = node.getNextSibling()) {
                if (!(node instanceof Element stop) || !"stop".equals(localName(stop))) {
                    continue;
                }
                Element colorElement = directChild(stop, "color");
                Color color = color(workbook, colorElement);
                if (color == null) {
                    continue;
                }
                double position = parseDouble(stop.getAttribute("position"), 0);
                stops.add(new Stop(Units.clamp(position, 0, 1), color));
            }
            stops.sort(Comparator.comparingDouble(Stop::position));
            return new CellGradient(
                    "path".equalsIgnoreCase(gradient.getAttribute("type")),
                    parseDouble(gradient.getAttribute("degree"), 0),
                    parseDouble(gradient.getAttribute("left"), 0),
                    parseDouble(gradient.getAttribute("right"), 0),
                    parseDouble(gradient.getAttribute("top"), 0),
                    parseDouble(gradient.getAttribute("bottom"), 0),
                    stops);
        } catch (Exception e) {
            return null;
        }
    }

    Paint paint(Rectangle2D bounds) {
        List<Stop> normalized = normalizedStops();
        if (normalized.isEmpty()) {
            return null;
        }
        if (normalized.size() == 1) {
            return normalized.get(0).color();
        }
        float[] fractions = new float[normalized.size()];
        Color[] colors = new Color[normalized.size()];
        for (int i = 0; i < normalized.size(); i++) {
            fractions[i] = (float) normalized.get(i).position();
            colors[i] = normalized.get(i).color();
        }
        if (path) {
            return radialPaint(bounds, normalized, fractions, colors);
        }
        double radians = Math.toRadians(degree);
        double dx = Math.cos(radians);
        double dy = Math.sin(radians);
        double half = (Math.abs(bounds.getWidth() * dx)
                + Math.abs(bounds.getHeight() * dy)) / 2d;
        if (half < 0.001) {
            return colors[0];
        }
        double centerX = bounds.getCenterX();
        double centerY = bounds.getCenterY();
        Point2D start = new Point2D.Double(centerX - dx * half, centerY - dy * half);
        Point2D end = new Point2D.Double(centerX + dx * half, centerY + dy * half);
        return new LinearGradientPaint(start, end, fractions, colors);
    }

    private Paint radialPaint(Rectangle2D bounds, List<Stop> stops,
                              float[] fractions, Color[] colors) {
        double x1 = bounds.getX() + bounds.getWidth() * Units.clamp(left, 0, 1);
        double x2 = bounds.getX() + bounds.getWidth() * Units.clamp(right, 0, 1);
        double y1 = bounds.getY() + bounds.getHeight() * Units.clamp(top, 0, 1);
        double y2 = bounds.getY() + bounds.getHeight() * Units.clamp(bottom, 0, 1);
        double centerX = (x1 + x2) / 2d;
        double centerY = (y1 + y2) / 2d;
        double outer = Math.max(
                Math.hypot(centerX - bounds.getMinX(), centerY - bounds.getMinY()),
                Math.max(
                        Math.hypot(centerX - bounds.getMaxX(), centerY - bounds.getMinY()),
                        Math.max(
                                Math.hypot(centerX - bounds.getMinX(), centerY - bounds.getMaxY()),
                                Math.hypot(centerX - bounds.getMaxX(), centerY - bounds.getMaxY()))));
        outer = Math.max(1, outer);
        double inner = Math.min(Math.abs(x2 - x1), Math.abs(y2 - y1)) / 2d;
        double innerFraction = Units.clamp(inner / outer, 0, 0.95);
        if (innerFraction < 0.001) {
            return new RadialGradientPaint(new Point2D.Double(centerX, centerY),
                    (float) outer, fractions, colors);
        }
        float[] expanded = new float[colors.length + 1];
        Color[] expandedColors = new Color[colors.length + 1];
        expanded[0] = 0;
        expandedColors[0] = colors[0];
        for (int i = 0; i < colors.length; i++) {
            expanded[i + 1] = (float) (innerFraction
                    + stops.get(i).position() * (1 - innerFraction));
            expandedColors[i + 1] = colors[i];
        }
        return new RadialGradientPaint(new Point2D.Double(centerX, centerY),
                (float) outer, expanded, expandedColors);
    }

    private List<Stop> normalizedStops() {
        List<Stop> result = new ArrayList<>();
        for (Stop stop : stops) {
            if (!result.isEmpty()
                    && Math.abs(result.get(result.size() - 1).position() - stop.position()) < 0.0001) {
                result.set(result.size() - 1, stop);
            } else {
                result.add(stop);
            }
        }
        if (result.isEmpty()) {
            return result;
        }
        if (result.get(0).position() > 0.001) {
            result.add(0, new Stop(0, result.get(0).color()));
        }
        Stop last = result.get(result.size() - 1);
        if (last.position() < 0.999) {
            result.add(new Stop(1, last.color()));
        }
        return result;
    }

    private static Color color(XSSFWorkbook workbook, Element color) {
        if (color == null) {
            return null;
        }
        double tint = parseDouble(color.getAttribute("tint"), 0);
        Color base = null;
        String rgb = color.getAttribute("rgb");
        if (rgb != null && !rgb.isBlank()) {
            base = rgb(rgb);
        } else if (color.hasAttribute("theme")) {
            int index = (int) parseDouble(color.getAttribute("theme"), -1);
            XSSFColor themed = index < 0 || workbook.getTheme() == null
                    ? null : workbook.getTheme().getThemeColor(index);
            base = themed == null ? Color.BLACK : ColorUtil.toAwt(themed, workbook, Color.BLACK);
        } else if (color.hasAttribute("indexed")) {
            int index = (int) parseDouble(color.getAttribute("indexed"), -1);
            IndexedColorMap map = workbook.getStylesSource() == null
                    ? null : workbook.getStylesSource().getIndexedColors();
            byte[] bytes = index < 0 || map == null ? null : map.getRGB(index);
            if (bytes != null && bytes.length >= 3) {
                base = new Color(bytes[0] & 0xFF, bytes[1] & 0xFF, bytes[2] & 0xFF);
            }
        } else if ("1".equals(color.getAttribute("auto"))
                || "true".equalsIgnoreCase(color.getAttribute("auto"))) {
            base = Color.BLACK;
        }
        return ColorUtil.withTint(base, tint);
    }

    private static Color rgb(String value) {
        String hex = value.trim();
        if (hex.startsWith("#")) {
            hex = hex.substring(1);
        }
        try {
            long parsed = Long.parseLong(hex, 16);
            if (hex.length() <= 6) {
                return new Color((int) parsed);
            }
            return new Color(
                    (int) ((parsed >>> 16) & 0xFF),
                    (int) ((parsed >>> 8) & 0xFF),
                    (int) (parsed & 0xFF),
                    (int) ((parsed >>> 24) & 0xFF));
        } catch (NumberFormatException e) {
            return null;
        }
    }

    private static Element find(Element parent, String name) {
        if (name.equals(localName(parent))) {
            return parent;
        }
        for (Node node = parent.getFirstChild(); node != null; node = node.getNextSibling()) {
            if (node instanceof Element element) {
                Element found = find(element, name);
                if (found != null) {
                    return found;
                }
            }
        }
        return null;
    }

    private static Element directChild(Element parent, String name) {
        for (Node node = parent.getFirstChild(); node != null; node = node.getNextSibling()) {
            if (node instanceof Element element && name.equals(localName(element))) {
                return element;
            }
        }
        return null;
    }

    private static String localName(Element element) {
        String name = element.getLocalName();
        if (name != null) {
            return name;
        }
        String nodeName = element.getNodeName();
        int colon = nodeName.indexOf(':');
        return colon >= 0 ? nodeName.substring(colon + 1) : nodeName;
    }

    private static double parseDouble(String value, double fallback) {
        try {
            return value == null || value.isBlank()
                    ? fallback : Double.parseDouble(value.trim());
        } catch (NumberFormatException e) {
            return fallback;
        }
    }
}
