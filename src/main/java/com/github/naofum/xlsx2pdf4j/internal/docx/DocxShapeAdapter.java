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

package com.github.naofum.xlsx2pdf4j.internal.docx;

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
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;

import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.awt.Color;
import java.awt.Font;
import java.awt.geom.Rectangle2D;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.function.DoubleFunction;

/**
 * Adapts WordprocessingML {@code w:drawing} / {@code wps:wsp} shapes to the
 * shared DrawingML core. POI's XWPF shape surface is intentionally small, so
 * the DrawingML part is read from the paragraph's DOM.
 */
public final class DocxShapeAdapter {

    private static final double EMU_PER_POINT = 12700d;
    private static final double DEFAULT_INSET_LR = 91440d;
    private static final double DEFAULT_INSET_TB = 45720d;

    private DocxShapeAdapter() {
    }

    public static List<Element> drawings(XWPFParagraph paragraph) {
        if (paragraph == null || paragraph.getCTP() == null) {
            return List.of();
        }
        return drawingElements((Element) paragraph.getCTP().getDomNode());
    }

    public static List<Element> drawings(XWPFDocument document) {
        if (document == null || document.getDocument() == null) {
            return List.of();
        }
        return drawingElements((Element) document.getDocument().getDomNode());
    }

    private static List<Element> drawingElements(Element root) {
        List<Element> result = new ArrayList<>();
        for (Element drawing : descendants(root, "drawing")) {
            if (first(drawing, "inline") != null || first(drawing, "anchor") != null) {
                result.add(drawing);
            }
        }
        return result;
    }

    public static DocxShape fromDrawing(Element drawing, DrawingTheme theme,
                                        FontResolver fonts, double scale) {
        if (drawing == null) {
            return null;
        }
        Element inline = first(drawing, "inline");
        Element anchor = first(drawing, "anchor");
        Element container = inline != null ? inline : anchor;
        if (container == null) {
            return null;
        }
        Element wsp = firstDescendant(container, "wsp");
        if (wsp == null) {
            return null;
        }
        Element spPr = first(wsp, "spPr");
        Element prstGeom = first(spPr, "prstGeom");
        PresetShape preset = DocxPresetShapes.fromOoxml(attr(prstGeom, "prst"));
        ShapeAdjustments adjustments = adjustments(prstGeom);
        DrawingTheme colors = theme == null ? DrawingTheme.defaults() : theme;
        DrawingStyle style = DrawingStyle.parse(xml(wsp), colors.colors());

        double factor = scale <= 0 ? 1.0 : scale;
        Element extent = first(container, "extent");
        double width = emu(attr(extent, "cx")) / EMU_PER_POINT;
        double height = emu(attr(extent, "cy")) / EMU_PER_POINT;
        if (width <= 0 || height <= 0) {
            Element ext = firstDescendant(spPr, "ext");
            width = emu(attr(ext, "cx")) / EMU_PER_POINT;
            height = emu(attr(ext, "cy")) / EMU_PER_POINT;
        }
        double x = 0;
        double y = 0;
        if (inline == null) {
            x = positionOffset(container, "positionH") / EMU_PER_POINT;
            y = positionOffset(container, "positionV") / EMU_PER_POINT;
        }
        Rectangle2D frame = new Rectangle2D.Double(x * factor, y * factor,
                width * factor, height * factor);
        String name = attr(first(container, "docPr"), "name");
        ShapeTextBody text = textBody(wsp, preset, adjustments, colors, fonts,
                factor, width, height);
        double textWidth = width;
        double textHeight = height;
        double pageScale = factor;
        DoubleFunction<ShapeTextBody> textAtScale = additionalScale ->
                textBody(wsp, preset, adjustments, colors, fonts,
                        pageScale * additionalScale, textWidth, textHeight);
        return new DocxShape(name, preset, adjustments, style, frame, text,
                inline != null, textAtScale);
    }

    /**
     * Builds the neutral model consumed by
     * {@link com.github.naofum.xlsx2pdf4j.internal.drawingml.ShapePainter}.
     */
    public static DrawingShape drawingShape(DocxShape shape) {
        if (shape == null) {
            return null;
        }
        return new DrawingShape(shape.name(), shape.frame(), null, shape.preset(),
                shape.adjustments(), null, shape.style(), shape.text(), shape.textAtScale());
    }

    private static ShapeTextBody textBody(Element wsp, PresetShape preset,
                                          ShapeAdjustments adjustments,
                                          DrawingTheme theme, FontResolver fonts,
                                          double scale, double width, double height) {
        Element txbx = first(wsp, "txbx");
        Element content = first(txbx, "txbxContent");
        if (content == null) {
            return null;
        }
        Element bodyPr = first(wsp, "bodyPr");
        TextInsets presetInsets = preset == null ? TextInsets.NONE
                : PresetTextInsets.forShape(preset, width, height, adjustments);
        TextInsets insets = presetInsets.plus(bodyInsets(bodyPr));

        Font baseFont = fonts.resolve("Arial", 11 * scale, false, false);
        List<TextLayoutEngine.Paragraph> paragraphs = new ArrayList<>();
        List<TextLayoutEngine.StyledRun> runs = new ArrayList<>();
        StringBuilder allText = new StringBuilder();
        boolean hasText = false;

        List<Element> paragraphElements = children(content, "p");
        for (int i = 0; i < paragraphElements.size(); i++) {
            Element paragraph = paragraphElements.get(i);
            HorizontalAlignment alignment = paragraphAlignment(paragraph);
            List<TextLayoutEngine.StyledRun> paragraphRuns = new ArrayList<>();
            for (Element run : descendants(paragraph, "r")) {
                String text = runText(run);
                if (text.isEmpty()) {
                    continue;
                }
                Element rPr = first(run, "rPr");
                Element rFonts = first(rPr, "rFonts");
                String family = attr(rFonts, "ascii");
                if (family.isBlank()) {
                    family = attr(rFonts, "hAnsi");
                }
                if (family.isBlank()) {
                    family = attr(rFonts, "eastAsia");
                }
                if (family.isBlank()) {
                    family = "Arial";
                }
                double size = size(rPr);
                boolean bold = onOff(rPr, "b");
                boolean italic = onOff(rPr, "i");
                boolean underline = underline(rPr);
                boolean strike = onOff(rPr, "strike") || onOff(rPr, "dstrike");
                Color color = color(rPr, theme);
                Font font = fonts.resolveForText(family, size * scale, bold, italic, text);
                paragraphRuns.add(new TextLayoutEngine.StyledRun(
                        text, font, color, underline, strike));
                allText.append(text);
                hasText = true;
            }
            Font fallback = paragraphRuns.isEmpty() ? baseFont : paragraphRuns.get(0).font();
            paragraphs.add(new TextLayoutEngine.Paragraph(paragraphRuns, alignment, fallback));
            runs.addAll(paragraphRuns);
            if (i + 1 < paragraphElements.size()) {
                Font newlineFont = runs.isEmpty() ? baseFont : runs.get(0).font();
                Color newlineColor = runs.isEmpty()
                        ? theme.color("tx1") : runs.get(0).color();
                runs.add(new TextLayoutEngine.StyledRun("\n", newlineFont,
                        newlineColor == null ? Color.BLACK : newlineColor, false, false));
                allText.append('\n');
            }
        }
        if (!hasText || runs.isEmpty()) {
            return null;
        }

        Element normAutofit = first(bodyPr, "normAutofit");
        boolean fontScaleSet = normAutofit != null
                && !attr(normAutofit, "fontScale").isBlank();
        double lineSpacing = 1.0;
        if (normAutofit != null && !attr(normAutofit, "lnSpcReduction").isBlank()) {
            lineSpacing = 1.0 - number(attr(normAutofit, "lnSpcReduction")) / 100000d;
        }
        String vert = attr(bodyPr, "vert");
        boolean eastAsianVertical = "eaVert".equals(vert)
                || "wordArtVertRtl".equals(vert);
        return new ShapeTextBody(paragraphs, runs, allText.toString(), insets,
                vertical(attr(bodyPr, "anchor")), wrap(bodyPr), anchorCtr(bodyPr),
                eastAsianVertical, normAutofit != null, fontScaleSet, lineSpacing);
    }

    private static ShapeAdjustments adjustments(Element prstGeom) {
        return (name, fallback) -> {
            Element avLst = first(prstGeom, "avLst");
            if (avLst == null) {
                return fallback;
            }
            for (Element guide : children(avLst, "gd")) {
                if (!name.equals(attr(guide, "name"))) {
                    continue;
                }
                String formula = attr(guide, "fmla").trim();
                if (formula.startsWith("val ")) {
                    try {
                        return Double.parseDouble(formula.substring(4).trim());
                    } catch (NumberFormatException ignored) {
                        return fallback;
                    }
                }
                return fallback;
            }
            return fallback;
        };
    }

    private static TextInsets bodyInsets(Element bodyPr) {
        return new TextInsets(
                inset(bodyPr, "lIns", DEFAULT_INSET_LR),
                inset(bodyPr, "tIns", DEFAULT_INSET_TB),
                inset(bodyPr, "rIns", DEFAULT_INSET_LR),
                inset(bodyPr, "bIns", DEFAULT_INSET_TB));
    }

    private static double inset(Element bodyPr, String name, double fallbackEmu) {
        String value = attr(bodyPr, name);
        return (value.isBlank() ? fallbackEmu : emu(value)) / EMU_PER_POINT;
    }

    private static double positionOffset(Element container, String name) {
        Element position = first(container, name);
        String value = text(first(position, "posOffset"));
        return value.isBlank() ? 0 : emu(value);
    }

    private static double size(Element rPr) {
        String value = attr(first(rPr, "sz"), "val");
        if (value.isBlank()) {
            return 11;
        }
        try {
            return Double.parseDouble(value) / 2d;
        } catch (NumberFormatException e) {
            return 11;
        }
    }

    private static Color color(Element rPr, DrawingTheme theme) {
        Element color = first(rPr, "color");
        if (color == null) {
            Color fallback = theme.color("tx1");
            return fallback == null ? Color.BLACK : fallback;
        }
        String themeName = attr(color, "themeColor");
        if (!themeName.isBlank()) {
            Color themed = theme.color(themeColorName(themeName));
            if (themed != null) {
                return themed;
            }
        }
        String value = attr(color, "val");
        if (value.length() == 6 && !"auto".equalsIgnoreCase(value)) {
            try {
                return new Color(Integer.parseInt(value, 16));
            } catch (NumberFormatException ignored) {
                // Fall through to the default text color.
            }
        }
        Color fallback = theme.color("tx1");
        return fallback == null ? Color.BLACK : fallback;
    }

    private static String themeColorName(String name) {
        return switch (name) {
            case "dark1", "text1" -> "tx1";
            case "light1", "background1" -> "bg1";
            case "dark2", "text2" -> "tx2";
            case "light2", "background2" -> "bg2";
            case "hyperlink" -> "hlink";
            case "followedHyperlink" -> "folHlink";
            default -> name;
        };
    }

    private static HorizontalAlignment paragraphAlignment(Element paragraph) {
        String value = attr(first(first(paragraph, "pPr"), "jc"), "val");
        return switch (value) {
            case "center" -> HorizontalAlignment.CENTER;
            case "right", "end" -> HorizontalAlignment.RIGHT;
            case "both", "distribute" -> HorizontalAlignment.JUSTIFY;
            default -> HorizontalAlignment.LEFT;
        };
    }

    private static VerticalAlignment vertical(String anchor) {
        return switch (anchor) {
            case "ctr", "just", "dist" -> VerticalAlignment.CENTER;
            case "b" -> VerticalAlignment.BOTTOM;
            default -> VerticalAlignment.TOP;
        };
    }

    private static boolean wrap(Element bodyPr) {
        String value = attr(bodyPr, "wrap");
        return value.isBlank() || !"none".equals(value);
    }

    private static boolean anchorCtr(Element bodyPr) {
        String value = attr(bodyPr, "anchorCtr");
        return "1".equals(value) || "true".equalsIgnoreCase(value);
    }

    private static String runText(Element run) {
        StringBuilder text = new StringBuilder();
        for (Node node = run.getFirstChild(); node != null; node = node.getNextSibling()) {
            if (!(node instanceof Element child)) {
                continue;
            }
            switch (local(child)) {
                case "t" -> text.append(text(child));
                case "tab" -> text.append('\t');
                case "br", "cr" -> text.append('\n');
                case "noBreakHyphen" -> text.append('-');
                default -> {
                    // rPr and nested drawing content are not text.
                }
            }
        }
        return text.toString();
    }

    private static boolean onOff(Element parent, String name) {
        Element element = first(parent, name);
        if (element == null) {
            return false;
        }
        String value = attr(element, "val");
        return value.isBlank() || (!"0".equals(value) && !"false".equalsIgnoreCase(value));
    }

    private static boolean underline(Element rPr) {
        Element underline = first(rPr, "u");
        if (underline == null) {
            return false;
        }
        String value = attr(underline, "val");
        return value.isBlank() || !"none".equals(value);
    }

    private static long emu(String value) {
        if (value == null || value.isBlank()) {
            return 0;
        }
        try {
            return Long.parseLong(value.trim());
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    private static double number(String value) {
        try {
            return value == null || value.isBlank() ? 0 : Double.parseDouble(value.trim());
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    private static String text(Element element) {
        return element == null ? "" : nodeText(element);
    }

    private static String nodeText(Node node) {
        StringBuilder text = new StringBuilder();
        for (Node child = node.getFirstChild(); child != null; child = child.getNextSibling()) {
            if (child.getNodeType() == Node.TEXT_NODE
                    || child.getNodeType() == Node.CDATA_SECTION_NODE) {
                text.append(child.getNodeValue());
            } else if (child.getNodeType() == Node.ELEMENT_NODE) {
                text.append(nodeText(child));
            }
        }
        return text.toString();
    }

    private static String xml(Element element) {
        if (element == null) {
            return "";
        }
        try {
            Transformer transformer = TransformerFactory.newInstance().newTransformer();
            transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");
            StringWriter writer = new StringWriter();
            transformer.transform(new DOMSource(element), new StreamResult(writer));
            return writer.toString();
        } catch (TransformerException ignored) {
            return "";
        }
    }

    private static String attr(Element element, String name) {
        if (element == null) {
            return "";
        }
        String value = element.getAttribute(name);
        if (value != null && !value.isEmpty()) {
            return value;
        }
        NamedNodeMap attributes = element.getAttributes();
        for (int i = 0; i < attributes.getLength(); i++) {
            Node attribute = attributes.item(i);
            String local = attribute.getLocalName();
            if (local == null) {
                String nodeName = attribute.getNodeName();
                int colon = nodeName.indexOf(':');
                local = colon >= 0 ? nodeName.substring(colon + 1) : nodeName;
            }
            if (name.equals(local)) {
                return attribute.getNodeValue();
            }
        }
        return "";
    }

    private static Element first(Element parent, String name) {
        if (parent == null) {
            return null;
        }
        for (Node node = parent.getFirstChild(); node != null; node = node.getNextSibling()) {
            if (node instanceof Element element && name.equals(local(element))) {
                return element;
            }
        }
        return null;
    }

    private static List<Element> children(Element parent, String name) {
        List<Element> result = new ArrayList<>();
        if (parent == null) {
            return result;
        }
        for (Node node = parent.getFirstChild(); node != null; node = node.getNextSibling()) {
            if (node instanceof Element element && name.equals(local(element))) {
                result.add(element);
            }
        }
        return result;
    }

    private static Element firstDescendant(Element parent, String name) {
        if (parent == null) {
            return null;
        }
        for (Node node = parent.getFirstChild(); node != null; node = node.getNextSibling()) {
            if (!(node instanceof Element element)) {
                continue;
            }
            if (name.equals(local(element))) {
                return element;
            }
            Element found = firstDescendant(element, name);
            if (found != null) {
                return found;
            }
        }
        return null;
    }

    private static List<Element> descendants(Element parent, String name) {
        List<Element> result = new ArrayList<>();
        descendants(parent, name, result);
        return result;
    }

    private static void descendants(Element parent, String name, List<Element> result) {
        if (parent == null) {
            return;
        }
        for (Node node = parent.getFirstChild(); node != null; node = node.getNextSibling()) {
            if (!(node instanceof Element element)) {
                continue;
            }
            if (name.equals(local(element))) {
                result.add(element);
            }
            descendants(element, name, result);
        }
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
