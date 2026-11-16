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

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilderFactory;
import java.awt.Color;
import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;

/**
 * Minimal DrawingML style parser for shapes, connectors and pictures.
 */
public final class DrawingStyle {

    private static final Map<String, Color> THEME = DrawingTheme.defaults().colors();

    public Color fill;
    public boolean fillNone;
    public Color line;
    public boolean lineNone;
    public double lineWidth = 0.75;
    public String dash;
    public String compound = "sng";
    public String headEnd;
    public String tailEnd;
    public double rotationDegrees;
    public boolean flipH;
    public boolean flipV;
    public List<GradientStop> gradientStops = List.of();
    public double gradientAngle;
    public boolean radialGradient;
    public Shadow shadow;

    public record GradientStop(double position, Color color) {
    }

    public record Shadow(double dx, double dy, double blur, Color color) {
    }

    public static DrawingStyle parse(String xml) {
        return parse(xml, THEME);
    }

    public static DrawingStyle parse(String xml, Map<String, Color> theme) {
        DrawingStyle style = new DrawingStyle();
        if (xml == null || xml.isBlank()) {
            return style;
        }
        Map<String, Color> colors = theme == null ? THEME : theme;
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            Document doc = factory.newDocumentBuilder().parse(
                    new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)));
            Element root = doc.getDocumentElement();
            String rootName = local(root);
            Element spPr = "spPr".equals(rootName) ? root : first(root, "spPr");
            if (spPr == null && "xml-fragment".equals(rootName)) {
                // parse() may be called with spPr.xmlText(), whose root is a
                // fragment containing the spPr children directly.
                spPr = root;
            }
            if (spPr == null) {
                return style;
            }
            Element xfrm = first(spPr, "xfrm");
            if (xfrm != null) {
                String rot = xfrm.getAttribute("rot");
                if (!rot.isBlank()) {
                    style.rotationDegrees = Double.parseDouble(rot) / 60000d;
                }
                style.flipH = "1".equals(xfrm.getAttribute("flipH")) || "true".equalsIgnoreCase(xfrm.getAttribute("flipH"));
                style.flipV = "1".equals(xfrm.getAttribute("flipV")) || "true".equalsIgnoreCase(xfrm.getAttribute("flipV"));
            }
            Element shapeStyle = first(root, "style");
            Element noFill = first(spPr, "noFill");
            Element solidFill = first(spPr, "solidFill");
            Element gradFill = first(spPr, "gradFill");
            if (noFill != null) {
                style.fillNone = true;
            } else if (solidFill != null) {
                style.fill = colorFrom(solidFill, colors);
            } else if (gradFill != null) {
                parseGradient(style, gradFill, colors);
            } else {
                applyFillRef(style, shapeStyle, colors);
            }
            // The shape style provides the default line color. An a:ln without
            // an explicit fill inherits it; Excel writes this pattern for the
            // arrows in flowcharts.
            applyLineRef(style, shapeStyle, colors);
            Element ln = first(spPr, "ln");
            if (ln != null) {
                if (first(ln, "noFill") != null) {
                    style.lineNone = true;
                } else {
                    Color explicit = colorFrom(ln, colors);
                    if (explicit != null) {
                        style.line = explicit;
                        style.lineNone = false;
                    }
                    String width = ln.getAttribute("w");
                    if (!width.isBlank()) {
                        style.lineWidth = emuToPoints(Long.parseLong(width));
                    }
                    Element dash = first(ln, "prstDash");
                    if (dash != null) {
                        style.dash = dash.getAttribute("val");
                    }
                    String compound = ln.getAttribute("cmpd");
                    if (!compound.isBlank()) {
                        style.compound = compound;
                    }
                    Element head = first(ln, "headEnd");
                    if (head != null) {
                        style.headEnd = arrowType(head.getAttribute("type"));
                    }
                    Element tail = first(ln, "tailEnd");
                    if (tail != null) {
                        style.tailEnd = arrowType(tail.getAttribute("type"));
                    }
                }
            }
            // Only the direct effectLst is rendered. Effects stored in
            // a14:hiddenEffects are intentionally hidden by Excel.
            Element effectList = first(spPr, "effectLst");
            Element outer = effectList == null ? null : first(effectList, "outerShdw");
            if (outer != null) {
                style.shadow = parseShadow(outer, colors);
            }
        } catch (Exception ignored) {
            // Keep defaults; the shape will still be rendered approximately.
        }
        return style;
    }

    private static void applyFillRef(DrawingStyle style, Element shapeStyle,
                                     Map<String, Color> theme) {
        if (shapeStyle == null) {
            return;
        }
        Element fillRef = first(shapeStyle, "fillRef");
        if (fillRef == null) {
            return;
        }
        String index = fillRef.getAttribute("idx");
        if (index == null || index.isBlank() || "0".equals(index)) {
            return;
        }
        Color color = colorFrom(fillRef, theme);
        if (color != null) {
            style.fill = color;
        }
    }

    private static void applyLineRef(DrawingStyle style, Element shapeStyle,
                                     Map<String, Color> theme) {
        if (shapeStyle == null) {
            return;
        }
        Element lineRef = first(shapeStyle, "lnRef");
        if (lineRef == null) {
            return;
        }
        String index = lineRef.getAttribute("idx");
        if (index == null || index.isBlank()) {
            return;
        }
        if ("0".equals(index)) {
            // lnRef idx=0 means no line. Excel writes this for text boxes.
            style.lineNone = true;
            return;
        }
        Color color = colorFrom(lineRef, theme);
        if (color != null) {
            style.line = color;
        }
        style.lineWidth = switch (index) {
            case "2" -> 2.0;
            case "3" -> 3.0;
            default -> 1.0;
        };
    }

    private static void parseGradient(DrawingStyle style, Element gradFill,
                                      Map<String, Color> theme) {
        Element list = first(gradFill, "gsLst");
        if (list == null) {
            return;
        }
        List<GradientStop> stops = new ArrayList<>();
        for (Node node = list.getFirstChild(); node != null; node = node.getNextSibling()) {
            if (!(node instanceof Element gs) || !"gs".equals(local(gs))) {
                continue;
            }
            String pos = gs.getAttribute("pos");
            double position;
            try {
                position = Double.parseDouble(pos) / 100000d;
            } catch (NumberFormatException e) {
                position = 0;
            }
            Color color = colorFrom(gs, theme);
            if (color != null) {
                stops.add(new GradientStop(clamp(position, 0, 1), color));
            }
        }
        stops.sort(Comparator.comparingDouble(GradientStop::position));
        style.gradientStops = List.copyOf(stops);
        if (!stops.isEmpty()) {
            style.fill = stops.get(0).color();
        }
        Element lin = first(gradFill, "lin");
        if (lin != null && !lin.getAttribute("ang").isBlank()) {
            try {
                style.gradientAngle = Double.parseDouble(lin.getAttribute("ang")) / 60000d;
            } catch (NumberFormatException ignored) {
                style.gradientAngle = 0;
            }
        }
        Element path = first(gradFill, "path");
        style.radialGradient = path != null;
    }

    private static Shadow parseShadow(Element outer, Map<String, Color> theme) {
        double distance = parseDouble(outer.getAttribute("dist"), 0);
        double direction = parseDouble(outer.getAttribute("dir"), 0);
        double blur = parseDouble(outer.getAttribute("blurRad"), 0);
        double radians = Math.toRadians(direction / 60000d);
        double dx = emuToPoints((long) Math.round(distance * Math.cos(radians)));
        double dy = emuToPoints((long) Math.round(distance * Math.sin(radians)));
        Color color = colorFrom(outer, theme);
        if (color == null) {
            color = new Color(0, 0, 0, 102);
        } else if (color.getAlpha() == 255) {
            // Excel shadows are usually translucent even when the color has no
            // explicit alpha; do not paint a fully opaque copy.
            color = new Color(color.getRed(), color.getGreen(), color.getBlue(), 110);
        }
        return new Shadow(dx, dy, emuToPoints((long) blur), color);
    }

    private static String arrowType(String value) {
        if (value == null || value.isBlank() || "none".equals(value)) {
            return null;
        }
        return value;
    }

    private static double parseDouble(String value, double fallback) {
        try {
            return value == null || value.isBlank() ? fallback : Double.parseDouble(value);
        } catch (NumberFormatException e) {
            return fallback;
        }
    }

    private static Color colorFrom(Element container, Map<String, Color> theme) {
        if (container == null) {
            return null;
        }
        Element srgb = first(container, "srgbClr");
        Element scheme = first(container, "schemeClr");
        if (srgb == null && scheme == null) {
            Element solid = first(container, "solidFill");
            if (solid != null) {
                srgb = first(solid, "srgbClr");
                scheme = first(solid, "schemeClr");
            }
        }
        return colorWithAlpha(srgb, scheme, theme);
    }

    private static Color colorWithAlpha(Element srgb, Element scheme,
                                        Map<String, Color> theme) {
        Color base = color(srgb, scheme, theme);
        if (base == null) {
            return null;
        }
        Element container = srgb != null ? srgb : scheme;
        if (container != null) {
            base = applyColorTransforms(base, container);
        }
        Element alpha = container == null ? null : first(container, "alpha");
        if (alpha == null) {
            return base;
        }
        try {
            double value = Double.parseDouble(alpha.getAttribute("val")) / 100000d;
            int a = (int) Math.round(clamp(value, 0, 1) * 255);
            return new Color(base.getRed(), base.getGreen(), base.getBlue(), a);
        } catch (NumberFormatException e) {
            return base;
        }
    }

    private static Color applyColorTransforms(Color color, Element container) {
        double[] hsl = rgbToHsl(color);
        boolean changed = false;
        Element tint = first(container, "tint");
        if (tint != null) {
            double value = colorPercent(tint);
            hsl[2] = hsl[2] + (1 - hsl[2]) * value;
            changed = true;
        }
        Element shade = first(container, "shade");
        if (shade != null) {
            hsl[2] = hsl[2] * (1 - colorPercent(shade));
            changed = true;
        }
        Element lumMod = first(container, "lumMod");
        if (lumMod != null) {
            hsl[2] = hsl[2] * colorPercent(lumMod);
            changed = true;
        }
        Element lumOff = first(container, "lumOff");
        if (lumOff != null) {
            hsl[2] = hsl[2] + colorPercent(lumOff);
            changed = true;
        }
        if (!changed) {
            return color;
        }
        return hslToColor(hsl[0], hsl[1], clamp(hsl[2], 0, 1), color.getAlpha());
    }

    private static double colorPercent(Element element) {
        try {
            return Double.parseDouble(element.getAttribute("val")) / 100000d;
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    private static double[] rgbToHsl(Color color) {
        double r = color.getRed() / 255d;
        double g = color.getGreen() / 255d;
        double b = color.getBlue() / 255d;
        double max = Math.max(r, Math.max(g, b));
        double min = Math.min(r, Math.min(g, b));
        double lightness = (max + min) / 2;
        double hue = 0;
        double saturation = 0;
        if (max != min) {
            double delta = max - min;
            saturation = lightness > 0.5
                    ? delta / (2 - max - min)
                    : delta / (max + min);
            if (max == r) {
                hue = (g - b) / delta + (g < b ? 6 : 0);
            } else if (max == g) {
                hue = (b - r) / delta + 2;
            } else {
                hue = (r - g) / delta + 4;
            }
            hue /= 6;
        }
        return new double[]{hue, saturation, lightness};
    }

    private static Color hslToColor(double hue, double saturation, double lightness, int alpha) {
        double r;
        double g;
        double b;
        if (saturation == 0) {
            r = g = b = lightness;
        } else {
            double q = lightness < 0.5
                    ? lightness * (1 + saturation)
                    : lightness + saturation - lightness * saturation;
            double p = 2 * lightness - q;
            r = hueToRgb(p, q, hue + 1d / 3);
            g = hueToRgb(p, q, hue);
            b = hueToRgb(p, q, hue - 1d / 3);
        }
        return new Color(channel(r), channel(g), channel(b), alpha);
    }

    private static double hueToRgb(double p, double q, double t) {
        double value = t;
        if (value < 0) {
            value += 1;
        }
        if (value > 1) {
            value -= 1;
        }
        if (value < 1d / 6) {
            return p + (q - p) * 6 * value;
        }
        if (value < 1d / 2) {
            return q;
        }
        if (value < 2d / 3) {
            return p + (q - p) * (2d / 3 - value) * 6;
        }
        return p;
    }

    private static int channel(double value) {
        return (int) Math.round(clamp(value, 0, 1) * 255);
    }

    private static Color color(Element srgb, Element scheme, Map<String, Color> theme) {
        if (srgb != null) {
            String value = srgb.getAttribute("val");
            if (value != null && value.length() == 6) {
                try {
                    return new Color(Integer.parseInt(value, 16));
                } catch (NumberFormatException ignored) {
                    return null;
                }
            }
        }
        if (scheme != null) {
            return theme.get(scheme.getAttribute("val"));
        }
        return null;
    }

    private static Element first(Element parent, String name) {
        NodeList children = parent.getChildNodes();
        for (int i = 0; i < children.getLength(); i++) {
            Node node = children.item(i);
            if (node instanceof Element e && name.equals(local(e))) {
                return e;
            }
        }
        return null;
    }

    private static String local(Element element) {
        String name = element.getLocalName();
        if (name != null) {
            return name;
        }
        String nodeName = element.getNodeName();
        int colon = nodeName.indexOf(':');
        return colon >= 0 ? nodeName.substring(colon + 1) : nodeName;
    }

    private static double emuToPoints(long emu) {
        return emu / 12700d;
    }

    private static double clamp(double value, double min, double max) {
        return Math.max(min, Math.min(max, value));
    }
}
