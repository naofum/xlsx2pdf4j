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

import org.w3c.dom.Element;
import org.w3c.dom.Node;

import java.awt.geom.Arc2D;
import java.awt.geom.Path2D;
import java.awt.geom.Point2D;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Parses DrawingML {@code a:custGeom}. Path-space coordinates are scaled to
 * the supplied host shape frame, and the common guide formulas are evaluated.
 */
public final class CustomGeometryParser {

    private CustomGeometryParser() {
    }

    public static CustomGeometry parse(Element root, double frameX, double frameY,
                                       double frameWidth, double frameHeight) {
        if (root == null) {
            return null;
        }
        Element pathList = directChild(root, "pathLst");
        if (pathList == null) {
            return null;
        }
        List<CustomPath> paths = new ArrayList<>();
        for (Node node = pathList.getFirstChild(); node != null; node = node.getNextSibling()) {
            if (!(node instanceof Element pathElement) || !"path".equals(localName(pathElement))) {
                continue;
            }
            double pathWidth = number(pathElement.getAttribute("w"), 21600);
            double pathHeight = number(pathElement.getAttribute("h"), 21600);
            if (pathWidth <= 0 || pathHeight <= 0) {
                pathWidth = Math.max(1, frameWidth);
                pathHeight = Math.max(1, frameHeight);
            }
            Map<String, String> formulas = customFormulas(root);
            Map<String, Double> builtins = customBuiltins(pathWidth, pathHeight);
            Path2D path = new Path2D.Double();
            boolean hasPoint = false;
            for (Node child = pathElement.getFirstChild(); child != null; child = child.getNextSibling()) {
                if (!(child instanceof Element command)) {
                    continue;
                }
                switch (localName(command)) {
                    case "moveTo" -> {
                        Point2D point = commandPoint(command, 0,
                                frameX, frameY, frameWidth, frameHeight,
                                pathWidth, pathHeight, formulas, builtins);
                        if (point != null) {
                            path.moveTo(point.getX(), point.getY());
                            hasPoint = true;
                        }
                    }
                    case "lnTo" -> {
                        Point2D point = commandPoint(command, 0,
                                frameX, frameY, frameWidth, frameHeight,
                                pathWidth, pathHeight, formulas, builtins);
                        if (point != null && hasPoint) {
                            path.lineTo(point.getX(), point.getY());
                        }
                    }
                    case "cubicBezTo" -> {
                        Point2D c1 = commandPoint(command, 0,
                                frameX, frameY, frameWidth, frameHeight,
                                pathWidth, pathHeight, formulas, builtins);
                        Point2D c2 = commandPoint(command, 1,
                                frameX, frameY, frameWidth, frameHeight,
                                pathWidth, pathHeight, formulas, builtins);
                        Point2D end = commandPoint(command, 2,
                                frameX, frameY, frameWidth, frameHeight,
                                pathWidth, pathHeight, formulas, builtins);
                        if (c1 != null && c2 != null && end != null && hasPoint) {
                            path.curveTo(c1.getX(), c1.getY(), c2.getX(), c2.getY(),
                                    end.getX(), end.getY());
                        }
                    }
                    case "quadBezTo" -> {
                        Point2D c = commandPoint(command, 0,
                                frameX, frameY, frameWidth, frameHeight,
                                pathWidth, pathHeight, formulas, builtins);
                        Point2D end = commandPoint(command, 1,
                                frameX, frameY, frameWidth, frameHeight,
                                pathWidth, pathHeight, formulas, builtins);
                        if (c != null && end != null && hasPoint) {
                            path.quadTo(c.getX(), c.getY(), end.getX(), end.getY());
                        }
                    }
                    case "arcTo" -> {
                        Point2D current = path.getCurrentPoint();
                        if (current == null || !hasPoint) {
                            continue;
                        }
                        double wR = resolveCoordinate(command.getAttribute("wR"),
                                formulas, builtins);
                        double hR = resolveCoordinate(command.getAttribute("hR"),
                                formulas, builtins);
                        double start = resolveCoordinate(command.getAttribute("stAng"),
                                formulas, builtins) / 60000d;
                        double sweep = resolveCoordinate(command.getAttribute("swAng"),
                                formulas, builtins) / 60000d;
                        double rx = wR * frameWidth / pathWidth;
                        double ry = hR * frameHeight / pathHeight;
                        if (rx <= 0 || ry <= 0) {
                            continue;
                        }
                        // DrawingML angles increase clockwise with y pointing
                        // down; Java2D uses the opposite sweep direction.
                        double angle = Math.toRadians(start);
                        double centerX = current.getX() - rx * Math.cos(angle);
                        double centerY = current.getY() - ry * Math.sin(angle);
                        path.append(new Arc2D.Double(centerX - rx, centerY - ry,
                                2 * rx, 2 * ry, -start, -sweep, Arc2D.OPEN), true);
                    }
                    case "close" -> {
                        if (hasPoint) {
                            path.closePath();
                        }
                    }
                    default -> {
                        // Unsupported custGeom command; ignore.
                    }
                }
            }
            String fillMode = pathElement.getAttribute("fill");
            if (fillMode == null || fillMode.isBlank()) {
                fillMode = "norm";
            }
            boolean stroke = !"false".equalsIgnoreCase(pathElement.getAttribute("stroke"));
            if (hasPoint) {
                paths.add(new CustomPath(path, !"none".equals(fillMode), stroke, fillMode));
            }
        }
        return paths.isEmpty() ? null : new CustomGeometry(paths);
    }

    private static Point2D commandPoint(Element command, int index,
                                        double frameX, double frameY,
                                        double frameWidth, double frameHeight,
                                        double pathWidth, double pathHeight,
                                        Map<String, String> formulas,
                                        Map<String, Double> builtins) {
        Element point = nthChild(command, "pt", index);
        if (point == null) {
            return null;
        }
        double x = resolveCoordinate(point.getAttribute("x"), formulas, builtins);
        double y = resolveCoordinate(point.getAttribute("y"), formulas, builtins);
        return new Point2D.Double(
                frameX + x * frameWidth / pathWidth,
                frameY + y * frameHeight / pathHeight);
    }

    private static Map<String, String> customFormulas(Element root) {
        Map<String, String> formulas = new HashMap<>();
        for (String listName : new String[]{"avLst", "gdLst"}) {
            Element list = directChild(root, listName);
            if (list == null) {
                continue;
            }
            for (Node node = list.getFirstChild(); node != null; node = node.getNextSibling()) {
                if (node instanceof Element guide && "gd".equals(localName(guide))) {
                    formulas.put(guide.getAttribute("name"), guide.getAttribute("fmla"));
                }
            }
        }
        return formulas;
    }

    private static Map<String, Double> customBuiltins(double w, double h) {
        Map<String, Double> values = new HashMap<>();
        values.put("w", w);
        values.put("h", h);
        values.put("l", 0d);
        values.put("t", 0d);
        values.put("r", w);
        values.put("b", h);
        values.put("hc", w / 2d);
        values.put("vc", h / 2d);
        values.put("wd2", w / 2d);
        values.put("hd2", h / 2d);
        values.put("wd3", w / 3d);
        values.put("hd3", h / 3d);
        values.put("wd4", w / 4d);
        values.put("hd4", h / 4d);
        values.put("wd8", w / 8d);
        values.put("hd8", h / 8d);
        values.put("ss", Math.min(w, h));
        values.put("ls", Math.max(w, h));
        values.put("cd2", 10800000d);
        values.put("cd4", 5400000d);
        values.put("cd8", 2700000d);
        values.put("3cd4", 16200000d);
        values.put("3cd8", 8100000d);
        values.put("5cd8", 13500000d);
        values.put("7cd8", 18900000d);
        return values;
    }

    private static double resolveCoordinate(String value, Map<String, String> formulas,
                                            Map<String, Double> builtins) {
        if (value == null || value.isBlank()) {
            return 0;
        }
        try {
            return Double.parseDouble(value.trim());
        } catch (NumberFormatException e) {
            return evalGuide(value.trim(), formulas, builtins, new HashSet<>());
        }
    }

    private static double evalGuide(String name, Map<String, String> formulas,
                                    Map<String, Double> builtins, Set<String> visiting) {
        Double builtin = builtins.get(name);
        if (builtin != null) {
            return builtin;
        }
        String formula = formulas.get(name);
        if (formula == null || !visiting.add(name)) {
            return 0;
        }
        try {
            return evalFormula(formula, formulas, builtins, visiting);
        } finally {
            visiting.remove(name);
        }
    }

    private static double evalFormula(String formula, Map<String, String> formulas,
                                      Map<String, Double> builtins, Set<String> visiting) {
        String[] tokens = formula.trim().split("\\s+");
        String op = tokens.length == 0 ? "" : tokens[0];
        return switch (op) {
            case "val" -> token(tokens, 1, formulas, builtins, visiting);
            case "*/" -> divide(
                    token(tokens, 1, formulas, builtins, visiting)
                            * token(tokens, 2, formulas, builtins, visiting),
                    token(tokens, 3, formulas, builtins, visiting));
            case "+-" -> token(tokens, 1, formulas, builtins, visiting)
                    + token(tokens, 2, formulas, builtins, visiting)
                    - token(tokens, 3, formulas, builtins, visiting);
            case "+/" -> divide(
                    token(tokens, 1, formulas, builtins, visiting)
                            + token(tokens, 2, formulas, builtins, visiting),
                    token(tokens, 3, formulas, builtins, visiting));
            case "?:" -> token(tokens, 1, formulas, builtins, visiting) > 0
                    ? token(tokens, 2, formulas, builtins, visiting)
                    : token(tokens, 3, formulas, builtins, visiting);
            case "abs" -> Math.abs(token(tokens, 1, formulas, builtins, visiting));
            case "sqrt" -> Math.sqrt(Math.max(0,
                    token(tokens, 1, formulas, builtins, visiting)));
            case "min" -> Math.min(token(tokens, 1, formulas, builtins, visiting),
                    token(tokens, 2, formulas, builtins, visiting));
            case "max" -> Math.max(token(tokens, 1, formulas, builtins, visiting),
                    token(tokens, 2, formulas, builtins, visiting));
            case "pin" -> Math.max(token(tokens, 1, formulas, builtins, visiting),
                    Math.min(token(tokens, 2, formulas, builtins, visiting),
                            token(tokens, 3, formulas, builtins, visiting)));
            case "mod" -> Math.sqrt(
                    Math.pow(token(tokens, 1, formulas, builtins, visiting), 2)
                            + Math.pow(token(tokens, 2, formulas, builtins, visiting), 2)
                            + Math.pow(token(tokens, 3, formulas, builtins, visiting), 2));
            case "sin" -> token(tokens, 1, formulas, builtins, visiting)
                    * Math.sin(angleRadians(token(tokens, 2, formulas, builtins, visiting)));
            case "cos" -> token(tokens, 1, formulas, builtins, visiting)
                    * Math.cos(angleRadians(token(tokens, 2, formulas, builtins, visiting)));
            case "tan" -> token(tokens, 1, formulas, builtins, visiting)
                    * Math.tan(angleRadians(token(tokens, 2, formulas, builtins, visiting)));
            case "at2" -> Math.toDegrees(Math.atan2(
                    token(tokens, 2, formulas, builtins, visiting),
                    token(tokens, 1, formulas, builtins, visiting))) * 60000d;
            case "cat2", "sat2" -> {
                double radius = token(tokens, 1, formulas, builtins, visiting);
                double angle = Math.atan2(
                        token(tokens, 3, formulas, builtins, visiting),
                        token(tokens, 2, formulas, builtins, visiting));
                yield "cat2".equals(op) ? radius * Math.cos(angle) : radius * Math.sin(angle);
            }
            default -> 0d;
        };
    }

    private static double token(String[] tokens, int index, Map<String, String> formulas,
                                Map<String, Double> builtins, Set<String> visiting) {
        if (index < 0 || index >= tokens.length) {
            return 0;
        }
        String value = tokens[index];
        try {
            return Double.parseDouble(value);
        } catch (NumberFormatException e) {
            return evalGuide(value, formulas, builtins, visiting);
        }
    }

    private static double divide(double numerator, double denominator) {
        return denominator == 0 ? 0 : numerator / denominator;
    }

    private static double angleRadians(double sixtyThousandths) {
        return Math.toRadians(sixtyThousandths / 60000d);
    }

    private static Element directChild(Element parent, String name) {
        for (Node node = parent.getFirstChild(); node != null; node = node.getNextSibling()) {
            if (node instanceof Element element && name.equals(localName(element))) {
                return element;
            }
        }
        return null;
    }

    private static Element nthChild(Element parent, String name, int index) {
        int seen = 0;
        for (Node node = parent.getFirstChild(); node != null; node = node.getNextSibling()) {
            if (node instanceof Element element && name.equals(localName(element))) {
                if (seen++ == index) {
                    return element;
                }
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

    private static double number(String value, double fallback) {
        try {
            return value == null || value.isBlank() ? fallback : Double.parseDouble(value);
        } catch (NumberFormatException e) {
            return fallback;
        }
    }
}
