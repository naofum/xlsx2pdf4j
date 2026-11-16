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
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * Format-neutral DrawingML color scheme. Host adapters read their theme part
 * (for example {@code xl/theme/theme1.xml}) and pass it to
 * {@link #fromThemeXml(InputStream)}.
 */
public final class DrawingTheme {

    private final Map<String, Color> colors;

    private DrawingTheme(Map<String, Color> colors) {
        this.colors = Map.copyOf(colors);
    }

    public static DrawingTheme defaults() {
        Map<String, Color> colors = new HashMap<>();
        colors.put("dk1", Color.BLACK);
        colors.put("lt1", Color.WHITE);
        colors.put("dk2", new Color(0x1F, 0x49, 0x7D));
        colors.put("lt2", new Color(0xEE, 0xEC, 0xE1));
        colors.put("accent1", new Color(0x4F, 0x81, 0xBD));
        colors.put("accent2", new Color(0xC0, 0x50, 0x4D));
        colors.put("accent3", new Color(0x9B, 0xBB, 0x59));
        colors.put("accent4", new Color(0x80, 0x64, 0xA2));
        colors.put("accent5", new Color(0x4B, 0xAC, 0xC6));
        colors.put("accent6", new Color(0xF7, 0x96, 0x46));
        colors.put("hlink", new Color(0x00, 0x00, 0xFF));
        colors.put("folHlink", new Color(0x80, 0x00, 0x80));
        aliasCoreColors(colors);
        return new DrawingTheme(colors);
    }

    public static DrawingTheme fromThemeXml(InputStream input) {
        Map<String, Color> colors = new HashMap<>(defaults().colors);
        if (input == null) {
            return new DrawingTheme(colors);
        }
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            Document document = factory.newDocumentBuilder().parse(input);
            Element scheme = find(document.getDocumentElement(), "clrScheme");
            if (scheme != null) {
                for (Node node = scheme.getFirstChild(); node != null; node = node.getNextSibling()) {
                    if (!(node instanceof Element entry)) {
                        continue;
                    }
                    Color value = color(firstElement(entry));
                    if (value != null) {
                        colors.put(local(entry), value);
                    }
                }
                aliasCoreColors(colors);
            }
        } catch (Exception ignored) {
            // Fall back to the default Office palette.
        }
        return new DrawingTheme(colors);
    }

    public Color color(String name) {
        return colors.get(name);
    }

    public Map<String, Color> colors() {
        return colors;
    }

    private static void aliasCoreColors(Map<String, Color> colors) {
        colors.put("tx1", colors.getOrDefault("dk1", Color.BLACK));
        colors.put("bg1", colors.getOrDefault("lt1", Color.WHITE));
        colors.put("tx2", colors.getOrDefault("dk2", Color.BLACK));
        colors.put("bg2", colors.getOrDefault("lt2", Color.WHITE));
    }

    private static Color color(Element element) {
        if (element == null) {
            return null;
        }
        String name = local(element);
        if ("srgbClr".equals(name) || "sysClr".equals(name)) {
            String value = element.getAttribute("lastClr");
            if (value == null || value.isBlank()) {
                value = element.getAttribute("val");
            }
            if (value != null && value.length() == 6) {
                try {
                    return new Color(Integer.parseInt(value, 16));
                } catch (NumberFormatException ignored) {
                    return null;
                }
            }
            if ("windowText".equals(element.getAttribute("val"))) {
                return Color.BLACK;
            }
            if ("window".equals(element.getAttribute("val"))) {
                return Color.WHITE;
            }
        }
        return null;
    }

    private static Element find(Element parent, String name) {
        if (parent == null) {
            return null;
        }
        if (name.equals(local(parent))) {
            return parent;
        }
        NodeList nodes = parent.getElementsByTagNameNS("*", name);
        if (nodes.getLength() > 0 && nodes.item(0) instanceof Element element) {
            return element;
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

    private static Element firstElement(Element parent) {
        for (Node node = parent.getFirstChild(); node != null; node = node.getNextSibling()) {
            if (node instanceof Element element) {
                return element;
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
}
