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

import com.github.naofum.xlsx2pdf4j.api.ConversionWarning;
import com.github.naofum.xlsx2pdf4j.api.WarningCode;
import com.github.naofum.xlsx2pdf4j.internal.chart.ChartRenderer;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.TextAlign;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFChildAnchor;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFConnector;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFGraphicFrame;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFShapeGroup;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextParagraph;
import org.apache.poi.xssf.usermodel.XSSFTextRun;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.CustomGeometry;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.CustomGeometryParser;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingShape;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingStyle;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingTheme;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.PresetShape;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.PresetTextInsets;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.ShapePainter;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.ShapeTextBody;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.TextInsets;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.PresetGeometry;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGeomGuide;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGroupTransform2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualConnectorProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPresetGeometry2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSolidColorFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBodyProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraphProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTransform2D;
import org.openxmlformats.schemas.drawingml.x2006.main.STTextAnchoringType;
import org.openxmlformats.schemas.drawingml.x2006.main.STTextVerticalType;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTConnector;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

import javax.imageio.ImageIO;
import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.Shape;
import java.awt.geom.AffineTransform;
import java.awt.geom.Area;
import java.awt.geom.Ellipse2D;
import java.awt.geom.Path2D;
import java.awt.geom.Point2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.IdentityHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

/**
 * Draws common DrawingML shapes. Unsupported shapes are approximated with
 * their bounding box and reported.
 */
public final class ShapeRenderer {

    private final FontResolver fonts;
    private final TextLayoutEngine textEngine;
    private final ChartRenderer chartRenderer;
    private final Map<Object, BufferedImage> imageCache = new HashMap<>();
    private Map<String, Rect> shapeRects = Map.of();
    private Map<String, Integer> shapeTypes = Map.of();
    private Map<Long, XSSFChart> chartsByFrame = Map.of();
    private Object warnedGraphicFrameSheet;
    private final Set<String> warnedGraphicFrames = new HashSet<>();
    private XSSFWorkbook themeWorkbook;
    private DrawingTheme drawingTheme = DrawingTheme.defaults();

    public ShapeRenderer(FontResolver fonts, TextLayoutEngine textEngine) {
        this.fonts = fonts;
        this.textEngine = textEngine;
        this.chartRenderer = new ChartRenderer(fonts);
    }

    public void drawShapes(Graphics2D g2, PageLayout page, List<ConversionWarning> warnings,
                           List<PlacedText> textLayer) {
        XSSFDrawing drawing = page.sheet.sheet.getDrawingPatriarch();
        if (drawing == null) {
            return;
        }
        List<XSSFShape> shapes = new ArrayList<>(drawing.getShapes());
        Map<Long, XSSFChart> charts = new HashMap<>();
        for (XSSFChart chart : drawing.getCharts()) {
            XSSFGraphicFrame frame = chart.getGraphicFrame();
            if (frame != null) {
                charts.put(frame.getId(), chart);
            }
        }
        chartsByFrame = charts;
        if (themeWorkbook != page.sheet.workbook) {
            themeWorkbook = page.sheet.workbook;
            drawingTheme = ThemeColors.read(themeWorkbook);
        }
        if (warnedGraphicFrameSheet != page.sheet.sheet) {
            warnedGraphicFrameSheet = page.sheet.sheet;
            warnedGraphicFrames.clear();
        }
        shapeRects = collectShapeRects(page);
        Shape oldClip = g2.getClip();
        // Excel clips drawings to the printable content area. Without this, a
        // shape anchored just below a page break can also appear in the blank
        // space at the bottom of the previous page.
        g2.clip(new Rectangle2D.Double(page.contentLeft, page.contentTop,
                page.contentWidth, page.contentHeight));
        try {
            // Arrow lines often sit visually on top of groups/boxes in the
            // source diagram. Draw them after the other shapes so their heads
            // are not covered by a group fill.
            for (int pass = 0; pass < 2; pass++) {
                for (XSSFShape shape : shapes) {
                    if (isLineShape(shape) != (pass == 1)) {
                        continue;
                    }
                    Rect bounds = anchor(page, shape);
                    if (bounds != null && !intersectsPage(page, bounds)) {
                        continue;
                    }
                    try {
                        drawShape(g2, page, shape, warnings, 0, null, textLayer);
                    } catch (RuntimeException e) {
                        warnings.add(new ConversionWarning(WarningCode.SHAPE_UNSUPPORTED,
                                page.sheet.name, shape.getClass().getSimpleName(),
                                "failed to draw shape: " + e.getMessage()));
                    }
                }
            }
        } finally {
            g2.setClip(oldClip);
        }
    }

    private static boolean isLineShape(XSSFShape shape) {
        return shapeType(shape) == ShapeTypes.LINE;
    }

    /**
     * A shape that ends exactly on a page boundary belongs to the preceding
     * page. The PDF clip alone would still show its stroke on the next page.
     */
    private static boolean intersectsPage(PageLayout page, Rect rect) {
        double left = page.contentLeft;
        double top = page.contentTop;
        double right = left + page.contentWidth;
        double bottom = top + page.contentHeight;
        double eps = 0.01;
        boolean xOverlap = rect.width <= 0.01
                ? rect.x >= left - eps && rect.x <= right + eps
                : rect.x + rect.width > left + eps && rect.x < right - eps;
        boolean yOverlap = rect.height <= 0.01
                ? rect.y >= top - eps && rect.y <= bottom + eps
                : rect.y + rect.height > top + eps && rect.y < bottom - eps;
        return xOverlap && yOverlap;
    }

    /**
     * ShapeTypes.typeFromString equivalent that tolerates custom geometry.
     * POI's XSSFSimpleShape#getShapeType throws when a:prstGeom is absent,
     * which is exactly the case for a:custGeom shapes.
     */
    private static int shapeType(XSSFShape shape) {
        try {
            if (shape instanceof XSSFSimpleShape simple) {
                CTShapeProperties properties = simple.getCTShape() == null
                        ? null : simple.getCTShape().getSpPr();
                if (properties == null || properties.getPrstGeom() == null) {
                    return -1;
                }
                return simple.getShapeType();
            }
            if (shape instanceof XSSFConnector connector) {
                CTShapeProperties properties = connector.getCTConnector() == null
                        ? null : connector.getCTConnector().getSpPr();
                if (properties == null || properties.getPrstGeom() == null) {
                    return -1;
                }
                return connector.getShapeType();
            }
        } catch (RuntimeException e) {
            return -1;
        }
        return -1;
    }

    private void drawShape(Graphics2D g2, PageLayout page, XSSFShape shape,
                           List<ConversionWarning> warnings, int depth, Rect override,
                           List<PlacedText> textLayer) {
        if (depth > 8) {
            return;
        }
        if (shape instanceof XSSFPicture picture) {
            drawPicture(g2, page, picture, warnings, override);
            return;
        }
        if (shape instanceof XSSFConnector connector) {
            drawConnector(g2, page, connector, warnings, override);
            return;
        }
        if (shape instanceof XSSFShapeGroup group) {
            drawGroup(g2, page, group, warnings, depth, override, textLayer);
            return;
        }
        if (shape instanceof XSSFGraphicFrame frame) {
            XSSFChart chart = chartsByFrame.get(frame.getId());
            if (chart != null) {
                Rect rect = override != null ? override : anchor(page, frame);
                drawChart(g2, page, chart, rect, warnings, textLayer);
            } else {
                warnGraphicFrame(page, frame, warnings);
            }
            return;
        }
        if (shape instanceof XSSFSimpleShape simple) {
            drawSimple(g2, page, simple, warnings, override, textLayer);
            return;
        }
        warnings.add(new ConversionWarning(WarningCode.SHAPE_UNSUPPORTED,
                page.sheet.name, shape.getClass().getSimpleName(), "unsupported shape type"));
    }

    private void drawChart(Graphics2D g2, PageLayout page, XSSFChart chart, Rect rect,
                           List<ConversionWarning> warnings, List<PlacedText> textLayer) {
        if (rect == null || rect.width <= 0 || rect.height <= 0) {
            return;
        }
        Shape oldClip = g2.getClip();
        Rectangle2D chartRect = new Rectangle2D.Double(rect.x, rect.y, rect.width, rect.height);
        g2.clip(chartRect);
        try {
            if (!chartRenderer.draw(g2, chart, chartRect, drawingTheme, textLayer)) {
                warnGraphicFrame(page, chart.getGraphicFrame(), warnings);
            }
        } catch (RuntimeException e) {
            warnGraphicFrame(page, chart.getGraphicFrame(), warnings);
        } finally {
            g2.setClip(oldClip);
        }
    }

    private void warnGraphicFrame(PageLayout page, XSSFGraphicFrame frame,
                                  List<ConversionWarning> warnings) {
        String uri = "";
        var ct = frame.getCTGraphicalObjectFrame();
        if (ct != null && ct.getGraphic() != null && ct.getGraphic().getGraphicData() != null) {
            String value = ct.getGraphic().getGraphicData().getUri();
            uri = value == null ? "" : value;
        }
        WarningCode code;
        String message;
        if (uri.contains("/chart")) {
            code = WarningCode.CHART_SKIPPED;
            message = "chart skipped";
        } else if (uri.contains("/diagram")) {
            code = WarningCode.SMART_ART_SKIPPED;
            message = "SmartArt skipped";
        } else {
            code = WarningCode.SHAPE_UNSUPPORTED;
            message = "unsupported graphic frame";
        }
        if (!warnedGraphicFrames.add(String.valueOf(frame.getId()))) {
            return;
        }
        warnings.add(new ConversionWarning(code, page.sheet.name,
                frame.getShapeName(), message));
    }

    private void drawGroup(Graphics2D g2, PageLayout page, XSSFShapeGroup group,
                           List<ConversionWarning> warnings, int depth, Rect override,
                           List<PlacedText> textLayer) {
        Rect groupRect = override != null ? override : anchor(page, group);
        if (groupRect == null || groupRect.width <= 0 || groupRect.height <= 0) {
            return;
        }
        GroupTransform transform = groupTransform(group, groupRect);
        for (XSSFShape child : directChildren(group)) {
            try {
                Rect childRect = groupChildRect(page, child, transform);
                if (childRect == null) {
                    warnings.add(new ConversionWarning(WarningCode.SHAPE_APPROXIMATED,
                            page.sheet.name, child.getClass().getSimpleName(),
                            "group child has no usable anchor"));
                    continue;
                }
                drawShape(g2, page, child, warnings, depth + 1, childRect, textLayer);
            } catch (RuntimeException e) {
                warnings.add(new ConversionWarning(WarningCode.SHAPE_UNSUPPORTED,
                        page.sheet.name, child.getClass().getSimpleName(),
                        "failed to draw group child: " + e.getMessage()));
            }
        }
    }

    /**
     * Connectors may reference shapes by id through stCxn/endCxn. Their own
     * anchor is only a cached bounding box and is often stale or reversed, so
     * collect the absolute rectangle of every shape before drawing.
     */
    private Map<String, Rect> collectShapeRects(PageLayout page) {
        Map<String, Rect> rects = new HashMap<>();
        Map<String, Integer> types = new HashMap<>();
        XSSFDrawing drawing = page.sheet.sheet.getDrawingPatriarch();
        if (drawing != null) {
            for (XSSFShape shape : drawing.getShapes()) {
                collectShapeRects(page, shape, null, rects, types, 0);
            }
        }
        shapeTypes = types;
        return rects;
    }

    private void collectShapeRects(PageLayout page, XSSFShape shape, Rect override,
                                   Map<String, Rect> rects, Map<String, Integer> types, int depth) {
        if (depth > 8) {
            return;
        }
        Rect rect = override != null ? override : anchor(page, shape);
        String id = shapeId(shape);
        if (id != null) {
            if (rect != null) {
                rects.put(id, rect);
            }
            if (shape instanceof XSSFSimpleShape simple) {
                types.put(id, shapeType(simple));
            }
        }
        if (shape instanceof XSSFShapeGroup group && rect != null) {
            GroupTransform transform = groupTransform(group, rect);
            for (XSSFShape child : directChildren(group)) {
                collectShapeRects(page, child, groupChildRect(page, child, transform), rects, types, depth + 1);
            }
        }
    }

    private static String shapeId(XSSFShape shape) {
        try {
            if (shape instanceof XSSFSimpleShape simple) {
                return String.valueOf(simple.getCTShape().getNvSpPr().getCNvPr().getId());
            }
            if (shape instanceof XSSFConnector connector) {
                return String.valueOf(connector.getCTConnector().getNvCxnSpPr().getCNvPr().getId());
            }
            if (shape instanceof XSSFShapeGroup group) {
                return String.valueOf(group.getCTGroupShape().getNvGrpSpPr().getCNvPr().getId());
            }
            if (shape instanceof XSSFPicture picture) {
                return String.valueOf(picture.getCTPicture().getNvPicPr().getCNvPr().getId());
            }
        } catch (RuntimeException e) {
            return null;
        }
        return null;
    }

    private static GroupTransform groupTransform(XSSFShapeGroup group, Rect groupRect) {
        CTGroupTransform2D xfrm = group.getCTGroupShape().getGrpSpPr().getXfrm();
        double chOffX = 0;
        double chOffY = 0;
        double chExtX = groupRect.width;
        double chExtY = groupRect.height;
        if (xfrm != null) {
            if (xfrm.getChOff() != null) {
                chOffX = coordinate(xfrm.getChOff().getX());
                chOffY = coordinate(xfrm.getChOff().getY());
            }
            if (xfrm.getChExt() != null && xfrm.getChExt().getCx() > 0 && xfrm.getChExt().getCy() > 0) {
                chExtX = xfrm.getChExt().getCx();
                chExtY = xfrm.getChExt().getCy();
            }
        }
        return new GroupTransform(
                groupRect.x - chOffX * (groupRect.width / chExtX),
                groupRect.y - chOffY * (groupRect.height / chExtY),
                groupRect.width / chExtX,
                groupRect.height / chExtY);
    }

    private Rect groupChildRect(PageLayout page, XSSFShape child, GroupTransform transform) {
        if (child.getAnchor() instanceof XSSFChildAnchor childAnchor) {
            return transform.rect(childAnchor);
        }
        if (child.getAnchor() instanceof XSSFClientAnchor) {
            return anchor(page, child);
        }
        CTTransform2D childXfrm = childTransform(child);
        if (childXfrm != null && childXfrm.getOff() != null && childXfrm.getExt() != null) {
            double offX = coordinate(childXfrm.getOff().getX());
            double offY = coordinate(childXfrm.getOff().getY());
            return transform.rect(offX, offY,
                    offX + childXfrm.getExt().getCx(),
                    offY + childXfrm.getExt().getCy());
        }
        CTGroupTransform2D childGroupXfrm = childGroupTransform(child);
        if (childGroupXfrm != null && childGroupXfrm.getOff() != null && childGroupXfrm.getExt() != null) {
            double offX = coordinate(childGroupXfrm.getOff().getX());
            double offY = coordinate(childGroupXfrm.getOff().getY());
            return transform.rect(offX, offY,
                    offX + childGroupXfrm.getExt().getCx(),
                    offY + childGroupXfrm.getExt().getCy());
        }
        return null;
    }

    private void drawSimple(Graphics2D g2, PageLayout page, XSSFSimpleShape shape,
                            List<ConversionWarning> warnings, Rect override,
                            List<PlacedText> textLayer) {
        Rect rect = override != null ? override : anchor(page, shape);
        int shapeType = shapeType(shape);
        boolean lineShape = shapeType == ShapeTypes.LINE
                || shapeType == ShapeTypes.STRAIGHT_CONNECTOR_1;
        if (rect == null
                || (rect.width <= 0 && rect.height <= 0)
                || (!lineShape && (rect.width <= 0 || rect.height <= 0))) {
            return;
        }
        rect = autofitRect(g2, shape, rect, page);
        DrawingStyle style = DrawingStyle.parse(shape.getCTShape().xmlText(), drawingTheme.colors());
        CustomGeometry customGeometry = customGeometry(shape, rect);
        CurvedArrowGeometry curvedArrow = switch (shapeType) {
            case ShapeTypes.CURVED_LEFT_ARROW ->
                    curvedHorizontalArrow(rect, false);
            case ShapeTypes.CURVED_RIGHT_ARROW ->
                    curvedHorizontalArrow(rect, true);
            case ShapeTypes.CURVED_UP_ARROW ->
                    curvedVerticalArrow(rect, shape, true);
            case ShapeTypes.CURVED_DOWN_ARROW ->
                    curvedVerticalArrow(rect, shape, false);
            default -> null;
        };
        Shape geometry = customGeometry != null ? null
                : curvedArrow != null ? curvedArrow.body()
                : geometry(shapeType, rect, warnings, page, shape);
        if (customGeometry == null && geometry == null) {
            return;
        }
        boolean foldedCorner = shapeType == ShapeTypes.FOLDED_CORNER;
        boolean strokeOnly = isStrokeOnlyShape(shapeType);
        double foldedAdjust = foldedCorner ? adjustValue(shape, "adj", 16667) : 0;
        Shape foldedFlap = foldedCorner ? foldedCornerFlap(rect, foldedAdjust) : null;
        Shape foldedOutline = foldedCorner ? foldedCornerOutline(rect, foldedAdjust) : null;
        Shape multiDocumentOutline = shapeType == ShapeTypes.FLOW_CHART_MULTIDOCUMENT
                ? multiDocumentOutline(rect.x, rect.y, rect.width, rect.height) : null;
        Shape diskOutline = shapeType == ShapeTypes.FLOW_CHART_MAGNETIC_DISK
                ? diskOutline(rect.x, rect.y, rect.width, rect.height) : null;
        Shape calloutLeader = switch (shapeType) {
            case ShapeTypes.BORDER_CALLOUT_1, ShapeTypes.CALLOUT_1,
                    ShapeTypes.ACCENT_CALLOUT_1, ShapeTypes.ACCENT_BORDER_CALLOUT_1 ->
                    borderCalloutLeader(shape, rect, 1);
            case ShapeTypes.BORDER_CALLOUT_2, ShapeTypes.CALLOUT_2,
                    ShapeTypes.ACCENT_CALLOUT_2, ShapeTypes.ACCENT_BORDER_CALLOUT_2 ->
                    borderCalloutLeader(shape, rect, 2);
            case ShapeTypes.BORDER_CALLOUT_3, ShapeTypes.CALLOUT_3,
                    ShapeTypes.ACCENT_CALLOUT_3, ShapeTypes.ACCENT_BORDER_CALLOUT_3 ->
                    borderCalloutLeader(shape, rect, 3);
            default -> null;
        };
        List<PresetFill> presetFills = switch (shapeType) {
            case ShapeTypes.BEVEL -> bevelFills(rect, shape);
            case ShapeTypes.CAN -> canFills(rect, shape);
            case ShapeTypes.CUBE -> cubeFills(rect, shape);
            default -> List.of();
        };
        boolean lineMarker = isLineMarkerShape(shapeType);
        Shape detailShape = switch (shapeType) {
            case ShapeTypes.FLOW_CHART_SUMMING_JUNCTION -> flowChartSummingDetail(rect);
            case ShapeTypes.FLOW_CHART_OR -> flowChartOrDetail(rect);
            case ShapeTypes.FLOW_CHART_SORT -> flowChartSortDetail(rect);
            case ShapeTypes.FLOW_CHART_OFFLINE_STORAGE -> flowChartOfflineDetail(rect);
            case ShapeTypes.FLOW_CHART_INTERNAL_STORAGE -> flowChartInternalDetail(rect);
            case ShapeTypes.FLOW_CHART_MAGNETIC_DRUM -> flowChartMagneticDrumDetail(rect);
            case ShapeTypes.FLOW_CHART_PREDEFINED_PROCESS -> flowChartPredefinedProcessDetail(rect);
            case ShapeTypes.RIBBON -> ribbonDetail(rect, false);
            case ShapeTypes.RIBBON_2 -> ribbonDetail(rect, true);
            case ShapeTypes.ELLIPSE_RIBBON -> ellipseRibbonDetail(rect, false);
            case ShapeTypes.ELLIPSE_RIBBON_2 -> ellipseRibbonDetail(rect, true);
            case ShapeTypes.LEFT_RIGHT_RIBBON -> leftRightRibbonDetail(rect);
            case ShapeTypes.VERTICAL_SCROLL -> scrollDetail(rect, true);
            case ShapeTypes.HORIZONTAL_SCROLL -> scrollDetail(rect, false);
            case ShapeTypes.SMILEY_FACE -> smileyDetail(rect);
            case ShapeTypes.NO_SMOKING -> null;
            case ShapeTypes.CHART_PLUS, ShapeTypes.CHART_STAR, ShapeTypes.CHART_X ->
                    chartMarkerDetail(rect, shapeType);
            case ShapeTypes.BEVEL -> bevelDetail(rect, shape);
            case ShapeTypes.CUBE -> cubeDetail(rect, shape);
            case ShapeTypes.ACCENT_CALLOUT_1, ShapeTypes.ACCENT_CALLOUT_2,
                    ShapeTypes.ACCENT_CALLOUT_3, ShapeTypes.ACCENT_BORDER_CALLOUT_1,
                    ShapeTypes.ACCENT_BORDER_CALLOUT_2,
                    ShapeTypes.ACCENT_BORDER_CALLOUT_3 -> accentCalloutBar(rect);
            default -> null;
        };

        Shape shadowGeometry = geometry;
        if (foldedFlap != null || curvedArrow != null) {
            Area combined = new Area(geometry);
            if (foldedFlap != null) {
                combined.add(new Area(foldedFlap));
            }
            if (curvedArrow != null) {
                combined.add(new Area(curvedArrow.head()));
            }
            shadowGeometry = combined;
        }
        Shape outline = curvedArrow != null ? curvedArrow.outline()
                : foldedOutline != null ? foldedOutline
                : multiDocumentOutline != null ? multiDocumentOutline
                : diskOutline != null ? diskOutline : geometry;
        Shape extraFill = foldedFlap != null ? foldedFlap
                : curvedArrow != null ? curvedArrow.head() : null;
        List<ShapePainter.Fill> fills = new ArrayList<>(presetFills.size());
        for (PresetFill fill : presetFills) {
            fills.add(new ShapePainter.Fill(fill.shape(), fill.mode()));
        }
        List<Shape> detailOutlines = new ArrayList<>(2);
        if (calloutLeader != null) {
            detailOutlines.add(calloutLeader);
        }
        if (detailShape != null) {
            detailOutlines.add(detailShape);
        }
        ShapePainter.Extras extras = new ShapePainter.Extras(shadowGeometry, outline,
                extraFill, fills, detailOutlines, strokeOnly, lineMarker, lineShape);
        Rectangle2D frame = new Rectangle2D.Double(rect.x, rect.y, rect.width, rect.height);
        Rect textRect = rect;
        ShapeTextBody text = shapeTextBody(shape, textRect, page.scale, 1.0);
        DrawingShape drawing = new DrawingShape(shape.getShapeName(), frame, geometry,
                null, null, customGeometry, style, text,
                additionalScale -> shapeTextBody(shape, textRect, page.scale, additionalScale));
        ShapePainter.draw(g2, drawing, extras, textEngine, textLayer);
    }

    /**
     * Parses DrawingML a:custGeom. Coordinates are path-space values scaled to
     * the shape rectangle; common guide formulas are evaluated.
     */
    private CustomGeometry customGeometry(XSSFSimpleShape shape, Rect rect) {
        CTShapeProperties properties = shape.getCTShape() == null
                ? null : shape.getCTShape().getSpPr();
        if (properties == null || properties.getCustGeom() == null) {
            return null;
        }
        Element root = (Element) properties.getCustGeom().getDomNode();
        return CustomGeometryParser.parse(root, rect.x, rect.y, rect.width, rect.height);
    }


    private ShapeTextBody shapeTextBody(XSSFSimpleShape shape, Rect rect,
                                        double layoutScale, double additionalScale) {
        CTTextBodyProperties body = bodyProperties(shape);
        TextInsets insets = presetTextInsets(shape, rect).plus(bodyInsets(body));
        VerticalAlignment vertical = verticalAlignment(body);
        boolean normAutofit = body != null && body.isSetNormAutofit();
        boolean fontScaleSet = normAutofit && body.getNormAutofit().isSetFontScale();
        double lineSpacing = 1.0;
        if (normAutofit && body.getNormAutofit().isSetLnSpcReduction()) {
            lineSpacing = 1.0 - percent(body.getNormAutofit().getLnSpcReduction(), 0.0);
        }
        TextContent content = textContent(shape, layoutScale, additionalScale);
        if (content == null) {
            return null;
        }
        boolean anchorCtr = body != null && body.isSetAnchorCtr() && body.getAnchorCtr();
        return new ShapeTextBody(content.paragraphs(), content.runs(), content.text(),
                insets, vertical, shape.getWordWrap(), anchorCtr,
                isEastAsianVertical(body), normAutofit, fontScaleSet, lineSpacing);
    }


    private static double percent(Object value, double fallback) {
        if (value == null) {
            return fallback;
        }
        if (value instanceof Number number) {
            return Units.clamp(number.doubleValue() / 100000d, 0, 1);
        }
        String text = value.toString().trim();
        try {
            if (text.endsWith("%")) {
                return Units.clamp(Double.parseDouble(
                        text.substring(0, text.length() - 1)) / 100d, 0, 1);
            }
            return Units.clamp(Double.parseDouble(text) / 100000d, 0, 1);
        } catch (NumberFormatException e) {
            return fallback;
        }
    }


    private TextContent textContent(XSSFSimpleShape shape, double layoutScale,
                                    double additionalScale) {
        List<XSSFTextParagraph> paragraphs = shape.getTextParagraphs();
        if (paragraphs.isEmpty()) {
            return null;
        }
        RunStyle base = baseStyle(paragraphs);
        if (base == null) {
            return null;
        }
        Font baseFont = fonts.resolveForText(base.fontName(),
                base.fontSize() * layoutScale * additionalScale,
                base.bold(), base.italic(), "");
        List<TextLayoutEngine.Paragraph> layoutParagraphs = new ArrayList<>();
        List<TextLayoutEngine.StyledRun> runs = new ArrayList<>();
        StringBuilder allText = new StringBuilder();
        boolean hasText = false;
        for (int i = 0; i < paragraphs.size(); i++) {
            XSSFTextParagraph paragraph = paragraphs.get(i);
            CTTextCharacterProperties defaults = defaultRunProperties(paragraph);
            List<TextLayoutEngine.StyledRun> paragraphRuns = new ArrayList<>();
            for (XSSFTextRun run : paragraph.getTextRuns()) {
                String runText = run.getText();
                if (runText == null || runText.isEmpty()) {
                    continue;
                }
                paragraphRuns.add(styledRun(run, runText, defaults, base,
                        layoutScale, additionalScale));
                allText.append(runText);
                hasText = true;
            }
            Font fallback = paragraphRuns.isEmpty() ? baseFont : paragraphRuns.get(0).font();
            layoutParagraphs.add(new TextLayoutEngine.Paragraph(paragraphRuns,
                    toHorizontal(paragraph.getTextAlign()), fallback));
            runs.addAll(paragraphRuns);
            if (i + 1 < paragraphs.size()) {
                Font newlineFont = runs.isEmpty() ? baseFont : runs.get(0).font();
                Color newlineColor = runs.isEmpty() ? base.color() : runs.get(0).color();
                runs.add(new TextLayoutEngine.StyledRun("\n", newlineFont,
                        newlineColor, false, false));
                allText.append('\n');
            }
        }
        if (!hasText || runs.isEmpty()) {
            return null;
        }
        return new TextContent(layoutParagraphs, runs, allText.toString());
    }

    /**
     * Approximates Excel's spAutoFit: a shape grows to fit its text. Width is
     * left unchanged because spAutoFit is often used on labels that keep a
     * fixed printed position.
     */
    private Rect autofitRect(Graphics2D g2, XSSFSimpleShape shape, Rect rect, PageLayout page) {
        CTTextBodyProperties body = bodyProperties(shape);
        if (body == null || !body.isSetSpAutoFit()) {
            return rect;
        }
        ShapeTextBody textBody = shapeTextBody(shape, rect, page.scale, 1.0);
        if (textBody == null) {
            return rect;
        }
        double width = rect.width;
        double height = rect.height;
        double maxWidth = Math.max(rect.width, page.pageWidth);
        double maxHeight = Math.max(rect.height, page.pageHeight);
        boolean wrap = textBody.wrap();
        for (int i = 0; i < 3; i++) {
            Rect candidate = new Rect(rect.x, rect.y, width, height);
            TextInsets insets = presetTextInsets(shape, candidate).plus(bodyInsets(body));
            if (!wrap) {
                double neededWidth = textEngine.richTextWidth(g2, textBody.runs())
                        + insets.left() + insets.right();
                width = Math.min(maxWidth, Math.max(width, neededWidth));
            }
            double availableWidth = Math.max(1, width - insets.left() - insets.right());
            double textHeight = textEngine.richTextHeight(g2, textBody.runs(),
                    availableWidth, wrap);
            double neededHeight = textHeight + insets.top() + insets.bottom();
            height = Math.min(maxHeight, Math.max(height, neededHeight));
            if (width <= rect.width + 0.5 && height <= rect.height + 0.5) {
                break;
            }
        }
        if (width <= rect.width + 0.5 && height <= rect.height + 0.5) {
            return rect;
        }
        double delta = height - rect.height;
        double y = switch (verticalAlignment(body)) {
            case BOTTOM -> rect.y - delta;
            case CENTER -> rect.y - delta / 2;
            default -> rect.y;
        };
        return new Rect(rect.x, y, width, height);
    }

    private static boolean isEastAsianVertical(CTTextBodyProperties body) {
        if (body == null || !body.isSetVert()) {
            return false;
        }
        STTextVerticalType.Enum vert = body.getVert();
        return vert == STTextVerticalType.EA_VERT
                || vert == STTextVerticalType.WORD_ART_VERT_RTL;
    }

    private static CTTextBodyProperties bodyProperties(XSSFSimpleShape shape) {
        if (shape.getCTShape() == null || shape.getCTShape().getTxBody() == null) {
            return null;
        }
        return shape.getCTShape().getTxBody().getBodyPr();
    }

    /**
     * Text rectangle defined by the DrawingML preset geometry. The values are
     * derived from the ECMA-376 presetShapeDefinitions formulas. Unknown shapes
     * keep the full bounding box.
     */
    private static TextInsets presetTextInsets(XSSFSimpleShape shape, Rect rect) {
        PresetShape preset = XlsxPresetShapes.from(shapeType(shape));
        if (preset == null) {
            return TextInsets.NONE;
        }
        return PresetTextInsets.forShape(preset, rect.width, rect.height,
                (name, fallback) -> adjustValue(shape, name, fallback));
    }

    private static double adjustValue(XSSFSimpleShape shape, String name, double fallback) {
        CTShapeProperties properties = shape.getCTShape() == null
                ? null : shape.getCTShape().getSpPr();
        CTPresetGeometry2D geometry = properties == null ? null : properties.getPrstGeom();
        if (geometry == null || !geometry.isSetAvLst()) {
            return fallback;
        }
        for (CTGeomGuide guide : geometry.getAvLst().getGdList()) {
            if (name.equals(guide.getName()) && guide.getFmla() != null
                    && guide.getFmla().startsWith("val ")) {
                try {
                    return Double.parseDouble(guide.getFmla().substring(4).trim());
                } catch (NumberFormatException e) {
                    return fallback;
                }
            }
        }
        return fallback;
    }

    private static double clamp(double value, double min, double max) {
        return Math.max(min, Math.min(max, value));
    }

    private static Shape multiDocumentOutline(double x, double y, double w, double h) {
        return PresetGeometry.multiDocumentOutline(x, y, w, h);
    }

    private static Shape diskOutline(double x, double y, double w, double h) {
        return PresetGeometry.diskOutline(x, y, w, h);
    }

    private static Shape arrow(double x, double y, double w, double h, double angle) {
        return PresetGeometry.arrow(x, y, w, h, angle);
    }

    private static Shape flipVertical(Shape shape, double centerY) {
        return PresetGeometry.flipVertical(shape, centerY);
    }

    private static void arcTo(Path2D path, double rx, double ry,
                              double startDegrees, double sweepDegrees) {
        PresetGeometry.arcTo(path, rx, ry, startDegrees, sweepDegrees);
    }

    private static VerticalAlignment verticalAlignment(CTTextBodyProperties body) {
        if (body == null || !body.isSetAnchor()) {
            return VerticalAlignment.TOP;
        }
        STTextAnchoringType.Enum anchor = body.getAnchor();
        if (anchor == STTextAnchoringType.CTR) {
            return VerticalAlignment.CENTER;
        }
        if (anchor == STTextAnchoringType.B) {
            return VerticalAlignment.BOTTOM;
        }
        if (anchor == STTextAnchoringType.JUST || anchor == STTextAnchoringType.DIST) {
            return VerticalAlignment.CENTER;
        }
        return VerticalAlignment.TOP;
    }

    private static final double DEFAULT_TEXT_LR = 91440 / 12700.0;
    private static final double DEFAULT_TEXT_TB = 45720 / 12700.0;

    private static TextInsets bodyInsets(CTTextBodyProperties body) {
        if (body == null) {
            return new TextInsets(DEFAULT_TEXT_LR, DEFAULT_TEXT_TB,
                    DEFAULT_TEXT_LR, DEFAULT_TEXT_TB);
        }
        return new TextInsets(
                textInset(body.getLIns(), body.isSetLIns(), DEFAULT_TEXT_LR),
                textInset(body.getTIns(), body.isSetTIns(), DEFAULT_TEXT_TB),
                textInset(body.getRIns(), body.isSetRIns(), DEFAULT_TEXT_LR),
                textInset(body.getBIns(), body.isSetBIns(), DEFAULT_TEXT_TB));
    }

    private static double textInset(Object value, boolean set, double fallback) {
        if (!set || value == null) {
            return fallback;
        }
        if (value instanceof Number number) {
            return number.doubleValue() / 12700.0;
        }
        try {
            return Double.parseDouble(String.valueOf(value)) / 12700.0;
        } catch (NumberFormatException e) {
            return fallback;
        }
    }

    private RunStyle baseStyle(List<XSSFTextParagraph> paragraphs) {
        for (XSSFTextParagraph paragraph : paragraphs) {
            CTTextCharacterProperties defaults = defaultRunProperties(paragraph);
            for (XSSFTextRun run : paragraph.getTextRuns()) {
                String text = run.getText();
                if (text != null && !text.isEmpty()) {
                    return runStyle(run, defaults, new RunStyle("MS PGothic", 10, Color.BLACK,
                            false, false, false, false));
                }
            }
        }
        return null;
    }

    private TextLayoutEngine.StyledRun styledRun(XSSFTextRun run, String text,
                                                 CTTextCharacterProperties defaults, RunStyle base,
                                                 double layoutScale, double additionalScale) {
        RunStyle style = runStyle(run, defaults, base);
        Font font = fonts.resolveForText(style.fontName(),
                style.fontSize() * layoutScale * additionalScale,
                style.bold(), style.italic(), text);
        return new TextLayoutEngine.StyledRun(text, font, style.color(),
                style.underline(), style.strike());
    }

    private RunStyle runStyle(XSSFTextRun run, CTTextCharacterProperties defaults, RunStyle base) {
        CTTextCharacterProperties properties = run.getXmlObject().getRPr();
        double fontSize = run.getFontSize() > 0 ? run.getFontSize()
                : defaults != null && defaults.isSetSz() && defaults.getSz() > 0
                        ? defaults.getSz() / 100.0
                        : base.fontSize();
        String fontName = run.getFontFamily();
        if (fontName == null || fontName.isBlank()) {
            fontName = base.fontName();
        }
        Color defaultColor = textColor(defaults, base.color());
        Color color = textColor(properties, defaultColor);
        Boolean runBold = bold(properties);
        Boolean runItalic = italic(properties);
        Boolean runUnderline = underline(properties);
        Boolean runStrike = strike(properties);
        boolean bold = runBold != null ? runBold : bool(bold(defaults));
        boolean italic = runItalic != null ? runItalic : bool(italic(defaults));
        boolean underline = runUnderline != null ? runUnderline : bool(underline(defaults));
        boolean strike = runStrike != null ? runStrike : bool(strike(defaults));
        return new RunStyle(fontName, fontSize, color, bold, italic, underline, strike);
    }

    /**
     * XSSFTextRun.getFontColor() does not resolve all scheme colors correctly
     * (for example bg1 is returned as black). Read the run's solid fill and
     * resolve it against the workbook theme first.
     */
    private Color textColor(CTTextCharacterProperties properties, Color fallback) {
        if (properties == null || properties.getSolidFill() == null) {
            return fallback;
        }
        CTSolidColorFillProperties fill = properties.getSolidFill();
        if (fill.isSetSrgbClr()) {
            byte[] rgb = fill.getSrgbClr().getVal();
            if (rgb != null && rgb.length == 3) {
                return new Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
            }
        }
        if (fill.isSetSchemeClr()) {
            String name = String.valueOf(fill.getSchemeClr().getVal());
            Color color = drawingTheme.color(name);
            if (color != null) {
                return color;
            }
        }
        return fallback;
    }

    private static CTTextCharacterProperties defaultRunProperties(XSSFTextParagraph paragraph) {
        CTTextParagraphProperties properties = paragraph.getXmlObject().getPPr();
        return properties == null ? null : properties.getDefRPr();
    }

    private static Boolean bold(CTTextCharacterProperties properties) {
        return properties != null && properties.isSetB() ? properties.getB() : null;
    }

    private static Boolean italic(CTTextCharacterProperties properties) {
        return properties != null && properties.isSetI() ? properties.getI() : null;
    }

    private static Boolean underline(CTTextCharacterProperties properties) {
        if (properties == null || !properties.isSetU()) {
            return null;
        }
        return !"none".equals(String.valueOf(properties.getU()));
    }

    private static Boolean strike(CTTextCharacterProperties properties) {
        if (properties == null || !properties.isSetStrike()) {
            return null;
        }
        return !"noStrike".equals(String.valueOf(properties.getStrike()));
    }

    private static boolean bool(Boolean value) {
        return value != null && value;
    }


    private record RunStyle(String fontName, double fontSize, Color color,
                            boolean bold, boolean italic, boolean underline, boolean strike) {
    }

    private record TextContent(List<TextLayoutEngine.Paragraph> paragraphs,
                               List<TextLayoutEngine.StyledRun> runs,
                               String text) {
    }

    /**
     * POI's XSSFShapeGroup iterator yields all descendants, not only direct
     * children. Recursing over it draws nested groups repeatedly with the wrong
     * transform. Map the XML child nodes to their wrappers to keep only direct
     * children.
     */
    private static List<XSSFShape> directChildren(XSSFShapeGroup group) {
        Map<Node, XSSFShape> byNode = new IdentityHashMap<>();
        for (XSSFShape shape : group) {
            Node node = shapeNode(shape);
            if (node != null) {
                byNode.put(node, shape);
            }
        }
        List<XSSFShape> children = new ArrayList<>();
        Node parent = group.getCTGroupShape().getDomNode();
        for (Node node = parent.getFirstChild(); node != null; node = node.getNextSibling()) {
            if (node.getNodeType() != Node.ELEMENT_NODE) {
                continue;
            }
            XSSFShape shape = byNode.get(node);
            if (shape != null) {
                children.add(shape);
            }
        }
        if (children.isEmpty()) {
            for (XSSFShape shape : group) {
                children.add(shape);
            }
        }
        return children;
    }

    private static Node shapeNode(XSSFShape shape) {
        if (shape instanceof XSSFSimpleShape simple) {
            return simple.getCTShape().getDomNode();
        }
        if (shape instanceof XSSFShapeGroup group) {
            return group.getCTGroupShape().getDomNode();
        }
        if (shape instanceof XSSFConnector connector) {
            return connector.getCTConnector().getDomNode();
        }
        if (shape instanceof XSSFPicture picture) {
            return picture.getCTPicture().getDomNode();
        }
        return null;
    }

    private void drawConnector(Graphics2D g2, PageLayout page, XSSFConnector connector,
                               List<ConversionWarning> warnings, Rect override) {
        Rect rect = override != null ? override : anchor(page, connector);
        if (rect == null) {
            return;
        }
        DrawingStyle style = DrawingStyle.parse(connector.getCTConnector().xmlText(), drawingTheme.colors());
        if (style.lineNone) {
            return;
        }
        ConnectorEndpoints endpoints = connectorEndpoints(connector, rect, style);
        ConnectorRoute route = connectorRoute(connector, endpoints, style);
        Shape path = route.path();
        if (style.shadow != null) {
            Shape shifted = AffineTransform.getTranslateInstance(style.shadow.dx(), style.shadow.dy())
                    .createTransformedShape(path);
            g2.setColor(style.shadow.color());
            g2.setStroke(new BasicStroke((float) Math.max(1, style.lineWidth + style.shadow.blur())));
            g2.draw(shifted);
        }
        g2.setColor(style.line == null ? Color.BLACK : style.line);
        ShapePainter.drawOutline(g2, path, style);
        if (style.headEnd != null) {
            Point2D next = firstDistinct(route.points(), endpoints.start());
            ShapePainter.drawArrowHead(g2, endpoints.start().getX(), endpoints.start().getY(),
                    next.getX(), next.getY(), style.headEnd, style.lineWidth);
        }
        if (style.tailEnd != null) {
            Point2D previous = lastDistinctBeforeEnd(route.points());
            ShapePainter.drawArrowHead(g2, endpoints.end().getX(), endpoints.end().getY(),
                    previous.getX(), previous.getY(), style.tailEnd, style.lineWidth);
        }
    }

    /**
     * Resolves the connector's endpoints from stCxn/endCxn. Excel recalculates
     * these glue points from the connected shapes, while the connector's own
     * anchor is only a cached, often stale, bounding box. When no glue point
     * exists, use the anchor corner selected by flipH/flipV.
     */
    private ConnectorEndpoints connectorEndpoints(XSSFConnector connector, Rect rect,
                                                   DrawingStyle style) {
        Rect startShape = null;
        Rect endShape = null;
        Integer startType = null;
        Integer endType = null;
        int startSite = -1;
        int endSite = -1;
        CTConnector ct = connector.getCTConnector();
        CTNonVisualConnectorProperties properties = ct == null || ct.getNvCxnSpPr() == null
                ? null : ct.getNvCxnSpPr().getCNvCxnSpPr();
        if (properties != null) {
            if (properties.isSetStCxn()) {
                String id = String.valueOf(properties.getStCxn().getId());
                startSite = (int) properties.getStCxn().getIdx();
                startShape = shapeRects.get(id);
                startType = shapeTypes.get(id);
            }
            if (properties.isSetEndCxn()) {
                String id = String.valueOf(properties.getEndCxn().getId());
                endSite = (int) properties.getEndCxn().getIdx();
                endShape = shapeRects.get(id);
                endType = shapeTypes.get(id);
            }
        }

        Point2D start = null;
        Point2D end = null;
        if (startShape != null && startSite >= 0 && startSite <= 3) {
            start = sitePoint(startShape, startSite, null, startType);
        }
        if (endShape != null && endSite >= 0 && endSite <= 3) {
            end = sitePoint(endShape, endSite, null, endType);
        }
        if (start == null && startShape != null && startSite >= 0) {
            start = sitePoint(startShape, startSite, end, startType);
        }
        if (end == null && endShape != null && endSite >= 0) {
            end = sitePoint(endShape, endSite, start, endType);
        }

        Point2D[] preset = presetEndpoints(rect, style);
        Point2D bboxStart = preset[0];
        Point2D bboxEnd = preset[1];
        if (start == null && end == null) {
            start = bboxStart;
            end = bboxEnd;
        } else if (start == null) {
            start = farther(bboxStart, bboxEnd, end);
        } else if (end == null) {
            end = farther(bboxStart, bboxEnd, start);
        }
        if (start.distanceSq(end) < 1e-9) {
            end = new Point2D.Double(end.getX() + 0.01, end.getY());
        }
        return new ConnectorEndpoints(start, end, startSite, endSite);
    }

    private static Point2D edgePoint(Rect rect, int site) {
        return switch (site) {
            case 0 -> new Point2D.Double(rect.x + rect.width / 2, rect.y);
            case 1 -> new Point2D.Double(rect.x, rect.y + rect.height / 2);
            case 2 -> new Point2D.Double(rect.x + rect.width / 2, rect.y + rect.height);
            default -> new Point2D.Double(rect.x + rect.width, rect.y + rect.height / 2);
        };
    }

    /**
     * Preset-specific connection sites. The flowchart disk and connector have
     * more than the four rectangular sites; other extra sites fall back to the
     * closest bounding-box edge.
     */
    private static Point2D sitePoint(Rect rect, int site, Point2D opposite, Integer shapeType) {
        if (shapeType != null && shapeType == ShapeTypes.FLOW_CHART_MAGNETIC_DISK) {
            return switch (site) {
                case 0 -> new Point2D.Double(rect.x + rect.width / 2, rect.y + rect.height / 3);
                case 1 -> edgePoint(rect, 0);
                case 2 -> edgePoint(rect, 1);
                case 3 -> edgePoint(rect, 2);
                case 4 -> edgePoint(rect, 3);
                default -> nearestEdgePoint(rect, opposite);
            };
        }
        if (shapeType != null && shapeType == ShapeTypes.FLOW_CHART_CONNECTOR) {
            double insetX = rect.width * 0.15;
            double insetY = rect.height * 0.15;
            return switch (site) {
                case 0 -> edgePoint(rect, 0);
                case 1 -> new Point2D.Double(rect.x + insetX, rect.y + insetY);
                case 2 -> edgePoint(rect, 1);
                case 3 -> new Point2D.Double(rect.x + insetX, rect.y + rect.height - insetY);
                case 4 -> edgePoint(rect, 2);
                case 5 -> new Point2D.Double(rect.x + rect.width - insetX,
                        rect.y + rect.height - insetY);
                case 6 -> edgePoint(rect, 3);
                case 7 -> new Point2D.Double(rect.x + rect.width - insetX, rect.y + insetY);
                default -> nearestEdgePoint(rect, opposite);
            };
        }
        if (site <= 3) {
            return edgePoint(rect, site);
        }
        return nearestEdgePoint(rect, opposite);
    }

    private static Point2D nearestEdgePoint(Rect rect, Point2D opposite) {
        if (opposite == null) {
            return new Point2D.Double(rect.x + rect.width / 2, rect.y + rect.height / 2);
        }
        Point2D best = edgePoint(rect, 0);
        double bestDistance = best.distanceSq(opposite);
        for (int candidate = 1; candidate < 4; candidate++) {
            Point2D point = edgePoint(rect, candidate);
            double distance = point.distanceSq(opposite);
            if (distance < bestDistance) {
                best = point;
                bestDistance = distance;
            }
        }
        return best;
    }

    private static Point2D farther(Point2D first, Point2D second, Point2D from) {
        return first.distanceSq(from) >= second.distanceSq(from) ? first : second;
    }

    private ConnectorRoute connectorRoute(XSSFConnector connector, ConnectorEndpoints endpoints,
                                          DrawingStyle style) {
        List<Point2D> points = new ArrayList<>();
        points.add(endpoints.start());
        switch (shapeType(connector)) {
            case ShapeTypes.STRAIGHT_CONNECTOR_1 -> points.add(endpoints.end());
            case ShapeTypes.BENT_CONNECTOR_2 -> bent2(points, endpoints, style);
            case ShapeTypes.BENT_CONNECTOR_3 -> bent3(points, endpoints, connector);
            case ShapeTypes.BENT_CONNECTOR_4 -> bent4(points, endpoints, connector);
            default -> points.add(endpoints.end());
        }
        Path2D path = new Path2D.Double();
        path.moveTo(points.get(0).getX(), points.get(0).getY());
        for (int i = 1; i < points.size(); i++) {
            path.lineTo(points.get(i).getX(), points.get(i).getY());
        }
        return new ConnectorRoute(path, points);
    }

    private static void bent2(List<Point2D> points, ConnectorEndpoints endpoints,
                              DrawingStyle style) {
        Point2D start = endpoints.start();
        Point2D end = endpoints.end();
        boolean known = endpoints.startSite() >= 0 && endpoints.endSite() >= 0;
        boolean verticalStart = isVertical(endpoints.startSite());
        boolean verticalEnd = isVertical(endpoints.endSite());
        if (known && verticalStart && !verticalEnd) {
            points.add(new Point2D.Double(start.getX(), end.getY()));
        } else if (known && !verticalStart && verticalEnd) {
            points.add(new Point2D.Double(end.getX(), start.getY()));
        } else if (known && verticalStart) {
            double middleY = (start.getY() + end.getY()) / 2;
            points.add(new Point2D.Double(start.getX(), middleY));
            points.add(new Point2D.Double(end.getX(), middleY));
        } else if (known) {
            double middleX = (start.getX() + end.getX()) / 2;
            points.add(new Point2D.Double(middleX, start.getY()));
            points.add(new Point2D.Double(middleX, end.getY()));
        } else if (isQuarterTurn(style.rotationDegrees)) {
            // The unrotated preset is horizontal-then-vertical. A 90/270
            // degree rotation makes it vertical-then-horizontal.
            points.add(new Point2D.Double(start.getX(), end.getY()));
        } else {
            points.add(new Point2D.Double(end.getX(), start.getY()));
        }
        points.add(end);
    }

    private static boolean isQuarterTurn(double degrees) {
        double angle = ((degrees % 360) + 360) % 360;
        return Math.abs(angle - 90) < 0.01 || Math.abs(angle - 270) < 0.01;
    }

    private static void bent3(List<Point2D> points, ConnectorEndpoints endpoints,
                              XSSFConnector connector) {
        Point2D start = endpoints.start();
        Point2D end = endpoints.end();
        boolean known = endpoints.startSite() >= 0 && endpoints.endSite() >= 0;
        boolean verticalStart = isVertical(endpoints.startSite());
        boolean verticalEnd = isVertical(endpoints.endSite());
        if (known && verticalStart && verticalEnd) {
            double middleY = (start.getY() + end.getY()) / 2;
            points.add(new Point2D.Double(start.getX(), middleY));
            points.add(new Point2D.Double(end.getX(), middleY));
        } else if (known && !verticalStart && !verticalEnd) {
            double adjustment = clamp(adjustValue(connector, "adj1", 50000), 0, 100000) / 100000d;
            double x = start.getX() + (end.getX() - start.getX()) * adjustment;
            points.add(new Point2D.Double(x, start.getY()));
            points.add(new Point2D.Double(x, end.getY()));
        } else if (known && verticalStart) {
            points.add(new Point2D.Double(start.getX(), end.getY()));
        } else if (known) {
            points.add(new Point2D.Double(end.getX(), start.getY()));
        } else if (Math.abs(end.getX() - start.getX()) >= Math.abs(end.getY() - start.getY())) {
            points.add(new Point2D.Double(end.getX(), start.getY()));
        } else {
            points.add(new Point2D.Double(start.getX(), end.getY()));
        }
        points.add(end);
    }

    private static void bent4(List<Point2D> points, ConnectorEndpoints endpoints,
                              XSSFConnector connector) {
        Point2D start = endpoints.start();
        Point2D end = endpoints.end();
        boolean known = endpoints.startSite() >= 0 && endpoints.endSite() >= 0;
        boolean verticalStart = isVertical(endpoints.startSite());
        boolean verticalEnd = isVertical(endpoints.endSite());
        double adjustmentX = adjustValue(connector, "adj1", 50000) / 100000d;
        double adjustmentY = adjustValue(connector, "adj2", 50000) / 100000d;
        double x = start.getX() + (end.getX() - start.getX()) * adjustmentX;
        double y = start.getY() + (end.getY() - start.getY()) * adjustmentY;
        if (known && !verticalStart && verticalEnd) {
            points.add(new Point2D.Double(x, start.getY()));
            points.add(new Point2D.Double(x, y));
            points.add(new Point2D.Double(end.getX(), y));
        } else if (known && verticalStart && !verticalEnd) {
            points.add(new Point2D.Double(start.getX(), y));
            points.add(new Point2D.Double(x, y));
            points.add(new Point2D.Double(x, end.getY()));
        } else if (known && verticalStart) {
            points.add(new Point2D.Double(start.getX(), y));
            points.add(new Point2D.Double(x, y));
            points.add(new Point2D.Double(x, end.getY()));
        } else if (known) {
            points.add(new Point2D.Double(x, start.getY()));
            points.add(new Point2D.Double(x, y));
            points.add(new Point2D.Double(end.getX(), y));
        } else if (Math.abs(end.getX() - start.getX()) >= Math.abs(end.getY() - start.getY())) {
            points.add(new Point2D.Double(end.getX(), start.getY()));
        } else {
            points.add(new Point2D.Double(start.getX(), end.getY()));
        }
        points.add(end);
    }

    private static boolean isVertical(int site) {
        return site == 0 || site == 2;
    }

    private static Point2D firstDistinct(List<Point2D> points, Point2D origin) {
        for (int i = 1; i < points.size(); i++) {
            if (points.get(i).distanceSq(origin) > 1e-9) {
                return points.get(i);
            }
        }
        return origin;
    }

    private static Point2D lastDistinctBeforeEnd(List<Point2D> points) {
        Point2D end = points.get(points.size() - 1);
        for (int i = points.size() - 2; i >= 0; i--) {
            if (points.get(i).distanceSq(end) > 1e-9) {
                return points.get(i);
            }
        }
        return end;
    }

    /**
     * The anchor is the post-rotation bounding box. Reconstruct the
     * pre-rotation box, select the preset path's start/end corners with the
     * flips, then rotate them back into page space.
     */
    private static Point2D[] presetEndpoints(Rect rect, DrawingStyle style) {
        double angle = ((style.rotationDegrees % 360) + 360) % 360;
        boolean swapped = Math.abs(angle - 90) < 0.01 || Math.abs(angle - 270) < 0.01;
        double width = swapped ? rect.height : rect.width;
        double height = swapped ? rect.width : rect.height;
        double centerX = rect.x + rect.width / 2;
        double centerY = rect.y + rect.height / 2;
        double x = centerX - width / 2;
        double y = centerY - height / 2;
        Point2D start = new Point2D.Double(
                style.flipH ? x + width : x,
                style.flipV ? y + height : y);
        Point2D end = new Point2D.Double(
                style.flipH ? x : x + width,
                style.flipV ? y : y + height);
        if (angle != 0) {
            AffineTransform rotation = AffineTransform.getRotateInstance(
                    Math.toRadians(angle), centerX, centerY);
            start = rotation.transform(start, null);
            end = rotation.transform(end, null);
        }
        return new Point2D[]{start, end};
    }

    private static double adjustValue(XSSFConnector connector, String name, double fallback) {
        CTShapeProperties properties = connector.getCTConnector() == null
                ? null : connector.getCTConnector().getSpPr();
        CTPresetGeometry2D geometry = properties == null ? null : properties.getPrstGeom();
        if (geometry == null || !geometry.isSetAvLst()) {
            return fallback;
        }
        for (CTGeomGuide guide : geometry.getAvLst().getGdList()) {
            if (name.equals(guide.getName()) && guide.getFmla() != null
                    && guide.getFmla().startsWith("val ")) {
                try {
                    return Double.parseDouble(guide.getFmla().substring(4).trim());
                } catch (NumberFormatException e) {
                    return fallback;
                }
            }
        }
        return fallback;
    }

    private void drawPicture(Graphics2D g2, PageLayout page, XSSFPicture picture,
                             List<ConversionWarning> warnings, Rect override) {
        Rect rect = override != null ? override : anchor(page, picture);
        if (rect == null || rect.width <= 0 || rect.height <= 0) {
            return;
        }
        try {
            BufferedImage image = imageCache.get(picture.getPictureData());
            if (image == null) {
                byte[] data = picture.getPictureData().getData();
                image = ImageIO.read(new ByteArrayInputStream(data));
                if (image != null) {
                    imageCache.put(picture.getPictureData(), image);
                }
            }
            if (image == null) {
                warnings.add(new ConversionWarning(WarningCode.IMAGE_UNSUPPORTED,
                        page.sheet.name, picture.getShapeName(), "unsupported image format"));
                return;
            }
            g2.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
            g2.drawImage(image, (int) Math.round(rect.x), (int) Math.round(rect.y),
                    (int) Math.round(rect.width), (int) Math.round(rect.height), null);
        } catch (IOException | RuntimeException e) {
            warnings.add(new ConversionWarning(WarningCode.IMAGE_UNSUPPORTED,
                    page.sheet.name, picture.getShapeName(), "failed to read image: " + e.getMessage()));
        }
    }

    private Rect anchor(PageLayout page, XSSFShape shape) {
        if (!(shape.getAnchor() instanceof XSSFClientAnchor anchor)) {
            return null;
        }
        SheetLayout layout = page.sheet;
        double x1 = safeColumnX(layout, anchor.getCol1()) + Units.emuToPoints(anchor.getDx1());
        double y1 = safeRowY(layout, anchor.getRow1())
                + SheetLayout.renderRowOffset(Units.emuToPoints(anchor.getDy1()));
        double x2 = safeColumnX(layout, anchor.getCol2()) + Units.emuToPoints(anchor.getDx2());
        double y2 = safeRowY(layout, anchor.getRow2())
                + SheetLayout.renderRowOffset(Units.emuToPoints(anchor.getDy2()));
        double sx = page.contentLeft + (Math.min(x1, x2) - layout.columnX[page.firstColumn]) * page.scale;
        double sy = page.contentTop + (Math.min(y1, y2) - layout.renderRowY[page.firstRow]) * page.scale;
        double w = Math.abs(x2 - x1) * page.scale;
        double h = Math.abs(y2 - y1) * page.scale;
        // A one-cell anchor with zero extents still has a usable size in some files.
        if (w < 0.5 && h < 0.5) {
            return null;
        }
        return new Rect(sx, sy, w, h);
    }

    private double safeColumnX(SheetLayout layout, int column) {
        int c = Math.max(0, Math.min(column, layout.columnX.length - 1));
        return layout.columnX[c];
    }

    private double safeRowY(SheetLayout layout, int row) {
        int r = Math.max(0, Math.min(row, layout.renderRowY.length - 1));
        return layout.renderRowY[r];
    }

    private static boolean isLineMarkerShape(int shapeType) {
        return shapeType == ShapeTypes.CHART_PLUS
                || shapeType == ShapeTypes.CHART_STAR
                || shapeType == ShapeTypes.CHART_X;
    }

    private Shape geometry(int type, Rect r, List<ConversionWarning> warnings,
                           PageLayout page, XSSFShape shape) {
        PresetShape preset = XlsxPresetShapes.from(type);
        Shape result = preset == null ? null : PresetGeometry.create(preset,
                r.x, r.y, r.width, r.height,
                (name, fallback) -> shapeAdjust(shape, name, fallback));
        if (result != null) {
            return result;
        }
        warnings.add(new ConversionWarning(WarningCode.SHAPE_APPROXIMATED,
                page.sheet.name, shape.getShapeName(),
                "shape type " + type + " approximated by a rectangle"));
        return new Rectangle2D.Double(r.x, r.y, r.width, r.height);
    }

    private static boolean isStrokeOnlyShape(int shapeType) {
        return switch (shapeType) {
            case ShapeTypes.ARC, ShapeTypes.LEFT_BRACE, ShapeTypes.RIGHT_BRACE,
                    ShapeTypes.BRACE_PAIR, ShapeTypes.LEFT_BRACKET,
                    ShapeTypes.RIGHT_BRACKET, ShapeTypes.BRACKET_PAIR -> true;
            default -> false;
        };
    }


    private record CurvedArrowGeometry(Shape body, Shape head, Shape outline) {
    }

    /**
     * ECMA-376 curvedLeftArrow / curvedRightArrow. The arrow is a partial
     * large ellipse with a separate arrowhead path.
     */
    /**
     * Default curvedLeftArrow geometry, normalized from Excel's rendering of
     * the preset (adj1=adj2/2, adj3=adj2/2 defaults).
     */
    private CurvedArrowGeometry curvedHorizontalArrow(Rect rect, boolean right) {
        if (rect.width <= 0 || rect.height <= 0) {
            return null;
        }
        Path2D body = new Path2D.Double();
        body.moveTo(tx(rect, 1.029136), ty(rect, 0.691156));
        body.curveTo(tx(rect, 1.029136), ty(rect, 0.518255),
                tx(rect, 0.580995), ty(rect, 0.378181),
                tx(rect, 0.028409), ty(rect, 0.378279));
        body.lineTo(tx(rect, 0.028418), ty(rect, 0.128022));
        body.curveTo(tx(rect, 0.581004), ty(rect, 0.127924),
                tx(rect, 1.029145), ty(rect, 0.267999),
                tx(rect, 1.029064), ty(rect, 0.440693));
        body.lineTo(tx(rect, 1.029136), ty(rect, 0.691156));
        body.closePath();

        Path2D head = new Path2D.Double();
        head.moveTo(tx(rect, 0.028391), ty(rect, 0.878792));
        head.lineTo(tx(rect, 0.235518), ty(rect, 0.621701));
        head.lineTo(tx(rect, 0.235428), ty(rect, 0.746877));
        head.curveTo(tx(rect, 0.551871), ty(rect, 0.725993),
                tx(rect, 0.816201), ty(rect, 0.658601),
                tx(rect, 0.945516), ty(rect, 0.565871));
        head.curveTo(tx(rect, 1.166633), ty(rect, 0.724428),
                tx(rect, 0.935167), ty(rect, 0.908749),
                tx(rect, 0.428716), ty(rect, 0.977988));
        head.curveTo(tx(rect, 0.366505), ty(rect, 0.9863),
                tx(rect, 0.301955), ty(rect, 0.992918),
                tx(rect, 0.2355), ty(rect, 0.99734));
        head.lineTo(tx(rect, 0.23558), ty(rect, 1.122419));
        head.closePath();

        Path2D outline = new Path2D.Double();
        outline.moveTo(tx(rect, 1.029136), ty(rect, 0.691156));
        outline.curveTo(tx(rect, 1.029136), ty(rect, 0.518255),
                tx(rect, 0.580995), ty(rect, 0.378181),
                tx(rect, 0.028409), ty(rect, 0.378279));
        outline.lineTo(tx(rect, 0.028418), ty(rect, 0.128022));
        outline.curveTo(tx(rect, 0.581004), ty(rect, 0.127924),
                tx(rect, 1.029145), ty(rect, 0.267999),
                tx(rect, 1.029064), ty(rect, 0.440693));
        outline.lineTo(tx(rect, 1.029136), ty(rect, 0.691156));
        outline.curveTo(tx(rect, 1.029154), ty(rect, 0.839021),
                tx(rect, 0.698063), ty(rect, 0.966622),
                tx(rect, 0.2355), ty(rect, 0.99734));
        outline.lineTo(tx(rect, 0.23558), ty(rect, 1.122419));
        outline.lineTo(tx(rect, 0.028391), ty(rect, 0.878792));
        outline.lineTo(tx(rect, 0.235518), ty(rect, 0.621701));
        outline.lineTo(tx(rect, 0.235428), ty(rect, 0.746877));
        outline.curveTo(tx(rect, 0.551871), ty(rect, 0.725993),
                tx(rect, 0.816201), ty(rect, 0.658601),
                tx(rect, 0.945516), ty(rect, 0.565871));

        if (right) {
            AffineTransform mirror = new AffineTransform(-1, 0, 0, 1,
                    rect.x + rect.x + rect.width, 0);
            return new CurvedArrowGeometry(
                    mirror.createTransformedShape(body),
                    mirror.createTransformedShape(head),
                    mirror.createTransformedShape(outline));
        }
        return new CurvedArrowGeometry(body, head, outline);
    }

    private static double tx(Rect rect, double normalized) {
        return rect.x + normalized * rect.width;
    }

    private static double ty(Rect rect, double normalized) {
        return rect.y + normalized * rect.height;
    }


















    /**
     * ECMA-376 curvedUpArrow / curvedDownArrow. The band curves through the
     * shape and ends in a separate arrowhead, using adj1 (shaft thickness),
     * adj2 (head width) and adj3 (head length).
     */
    private CurvedArrowGeometry curvedVerticalArrow(Rect rect, XSSFShape shape,
                                                     boolean up) {
        double w = rect.width;
        double h = rect.height;
        double ss = Math.max(0.001, Math.min(w, h));
        double a1 = clamp(shapeAdjust(shape, "adj1", 25000), 0, 100000);
        double maxAdj2 = 50000d * w / ss;
        double a2 = clamp(shapeAdjust(shape, "adj2", 50000), 0, maxAdj2);
        double th = ss * a1 / 100000d;
        double aw = ss * a2 / 100000d;
        double wR = w / 2d - (th + aw) / 4d;
        double q7 = 2 * wR;
        if (wR <= 0.1 || q7 <= 0.1) {
            Path2D empty = new Path2D.Double();
            return new CurvedArrowGeometry(arrow(rect.x, rect.y, w, h, 0),
                    empty, empty);
        }
        double idy = Math.sqrt(Math.max(0, q7 * q7 - th * th)) * h / q7;
        double ah = ss * clamp(shapeAdjust(shape, "adj3", 25000), 0, 100000) / 100000d;
        double x3 = wR + th;
        double dx = Math.sqrt(Math.max(0, h * h - ah * ah)) * wR / h;
        double x5 = wR + dx;
        double x7 = x3 + dx;
        double dh = (aw - th) / 2d;
        double x4 = x5 - dh;
        double x8 = x7 + dh;
        double x6 = w - aw / 2d;
        double y1 = ah;
        double swAng = Math.toDegrees(Math.atan2(dx, ah));
        double ix = (wR + x3) / 2d;
        double iy = idy;
        double dang2 = Math.toDegrees(Math.atan2(th / 2d, idy));
        double swAng2 = dang2 - swAng;
        double stAng3 = 90 - swAng;
        double swAng3 = swAng + dang2;
        double stAng2 = 90 - dang2;
        double x = rect.x;
        double y = rect.y;

        Path2D body = new Path2D.Double();
        body.moveTo(x + x6, y);
        body.lineTo(x + x8, y + y1);
        body.lineTo(x + x7, y + y1);
        arcTo(body, wR, h, stAng3, swAng3);
        arcTo(body, wR, h, stAng2, swAng2);
        body.lineTo(x + x4, y + y1);
        body.closePath();

        Path2D head = new Path2D.Double();
        head.moveTo(x + wR, y + h);
        arcTo(head, wR, h, 90, 90);
        head.lineTo(x + th, y);
        arcTo(head, wR, h, 180, -90);
        head.closePath();

        Path2D outline = new Path2D.Double();
        outline.moveTo(x + ix, y + iy);
        arcTo(outline, wR, h, stAng2, swAng2);
        outline.lineTo(x + x4, y + y1);
        outline.lineTo(x + x6, y);
        outline.lineTo(x + x8, y + y1);
        outline.lineTo(x + x7, y + y1);
        arcTo(outline, wR, h, stAng3, swAng);
        outline.lineTo(x + wR, y + h);
        arcTo(outline, wR, h, 90, 90);
        outline.lineTo(x + th, y);
        arcTo(outline, wR, h, 180, -90);

        if (up) {
            return new CurvedArrowGeometry(body, head, outline);
        }
        double centerY = y + h / 2d;
        return new CurvedArrowGeometry(
                flipVertical(body, centerY),
                flipVertical(head, centerY),
                flipVertical(outline, centerY));
    }





    private static double shapeAdjust(XSSFShape shape, String name, double fallback) {
        return shape instanceof XSSFSimpleShape simple
                ? adjustValue(simple, name, fallback) : fallback;
    }




























    private Shape flowChartSummingDetail(Rect rect) {
        double idx = rect.width / 2d * Math.cos(Math.toRadians(45));
        double idy = rect.height / 2d * Math.sin(Math.toRadians(45));
        Path2D path = new Path2D.Double();
        path.moveTo(rect.x + rect.width / 2d - idx, rect.y + rect.height / 2d - idy);
        path.lineTo(rect.x + rect.width / 2d + idx, rect.y + rect.height / 2d + idy);
        path.moveTo(rect.x + rect.width / 2d + idx, rect.y + rect.height / 2d - idy);
        path.lineTo(rect.x + rect.width / 2d - idx, rect.y + rect.height / 2d + idy);
        return path;
    }

    private Shape flowChartOrDetail(Rect rect) {
        Path2D path = new Path2D.Double();
        path.moveTo(rect.x + rect.width / 2d, rect.y);
        path.lineTo(rect.x + rect.width / 2d, rect.y + rect.height);
        path.moveTo(rect.x, rect.y + rect.height / 2d);
        path.lineTo(rect.x + rect.width, rect.y + rect.height / 2d);
        return path;
    }

    private Shape flowChartSortDetail(Rect rect) {
        Path2D path = new Path2D.Double();
        path.moveTo(rect.x, rect.y + rect.height / 2d);
        path.lineTo(rect.x + rect.width, rect.y + rect.height / 2d);
        return path;
    }

    private Shape flowChartOfflineDetail(Rect rect) {
        Path2D path = new Path2D.Double();
        path.moveTo(rect.x + rect.width * 0.4, rect.y + rect.height * 0.8);
        path.lineTo(rect.x + rect.width * 0.6, rect.y + rect.height * 0.8);
        return path;
    }

    private Shape flowChartInternalDetail(Rect rect) {
        Path2D path = new Path2D.Double();
        path.moveTo(rect.x + rect.width / 8d, rect.y);
        path.lineTo(rect.x + rect.width / 8d, rect.y + rect.height);
        path.moveTo(rect.x, rect.y + rect.height / 8d);
        path.lineTo(rect.x + rect.width, rect.y + rect.height / 8d);
        return path;
    }

    private Shape flowChartPredefinedProcessDetail(Rect rect) {
        Path2D path = new Path2D.Double();
        path.moveTo(rect.x + rect.width / 8d, rect.y);
        path.lineTo(rect.x + rect.width / 8d, rect.y + rect.height);
        path.moveTo(rect.x + rect.width * 7d / 8d, rect.y);
        path.lineTo(rect.x + rect.width * 7d / 8d, rect.y + rect.height);
        return path;
    }

    private Shape flowChartMagneticDrumDetail(Rect rect) {
        Path2D path = new Path2D.Double();
        double rx = rect.width / 6d;
        double ry = rect.height / 2d;
        path.moveTo(rect.x + rect.width - rx, rect.y + rect.height);
        arcTo(path, rx, ry, 90, 180);
        return path;
    }























































    private Shape smileyDetail(Rect rect) {
        Path2D path = new Path2D.Double();
        double eye = Math.min(rect.width, rect.height) * 0.08;
        path.append(new Ellipse2D.Double(
                rect.x + rect.width * 0.32 - eye / 2,
                rect.y + rect.height * 0.35 - eye / 2, eye, eye), false);
        path.append(new Ellipse2D.Double(
                rect.x + rect.width * 0.68 - eye / 2,
                rect.y + rect.height * 0.35 - eye / 2, eye, eye), false);
        path.moveTo(rect.x + rect.width * 0.28, rect.y + rect.height * 0.62);
        path.quadTo(rect.x + rect.width * 0.5, rect.y + rect.height * 0.82,
                rect.x + rect.width * 0.72, rect.y + rect.height * 0.62);
        return path;
    }

    private Shape bevelDetail(Rect rect, XSSFShape shape) {
        double ss = Math.min(rect.width, rect.height);
        double a = clamp(shapeAdjust(shape, "adj", 12500), 0, 50000);
        double x1 = ss * a / 100000d;
        double x2 = rect.width - x1;
        double y2 = rect.height - x1;
        Path2D path = new Path2D.Double();
        path.moveTo(rect.x + x1, rect.y + x1);
        path.lineTo(rect.x + x2, rect.y + x1);
        path.lineTo(rect.x + x2, rect.y + y2);
        path.lineTo(rect.x + x1, rect.y + y2);
        path.closePath();
        path.moveTo(rect.x, rect.y);
        path.lineTo(rect.x + x1, rect.y + x1);
        path.moveTo(rect.x, rect.y + rect.height);
        path.lineTo(rect.x + x1, rect.y + y2);
        path.moveTo(rect.x + rect.width, rect.y);
        path.lineTo(rect.x + x2, rect.y + x1);
        path.moveTo(rect.x + rect.width, rect.y + rect.height);
        path.lineTo(rect.x + x2, rect.y + y2);
        return path;
    }

    private List<PresetFill> bevelFills(Rect rect, XSSFShape shape) {
        double ss = Math.min(rect.width, rect.height);
        double a = clamp(shapeAdjust(shape, "adj", 12500), 0, 50000);
        double x1 = ss * a / 100000d;
        double x2 = rect.width - x1;
        double y2 = rect.height - x1;
        double l = rect.x;
        double t = rect.y;
        double r = rect.x + rect.width;
        double b = rect.y + rect.height;

        Path2D top = new Path2D.Double();
        top.moveTo(l, t);
        top.lineTo(r, t);
        top.lineTo(l + x2, t + x1);
        top.lineTo(l + x1, t + x1);
        top.closePath();

        Path2D bottom = new Path2D.Double();
        bottom.moveTo(l, b);
        bottom.lineTo(l + x1, t + y2);
        bottom.lineTo(l + x2, t + y2);
        bottom.lineTo(r, b);
        bottom.closePath();

        Path2D left = new Path2D.Double();
        left.moveTo(l, t);
        left.lineTo(l + x1, t + x1);
        left.lineTo(l + x1, t + y2);
        left.lineTo(l, b);
        left.closePath();

        Path2D right = new Path2D.Double();
        right.moveTo(r, t);
        right.lineTo(r, b);
        right.lineTo(l + x2, t + y2);
        right.lineTo(l + x2, t + x1);
        right.closePath();

        return List.of(
                new PresetFill(top, "lightenLess"),
                new PresetFill(bottom, "darkenLess"),
                new PresetFill(left, "lighten"),
                new PresetFill(right, "darken"));
    }

    private Shape chartMarkerDetail(Rect rect, int shapeType) {
        double l = rect.x;
        double t = rect.y;
        double r = rect.x + rect.width;
        double b = rect.y + rect.height;
        double cx = rect.x + rect.width / 2d;
        double cy = rect.y + rect.height / 2d;
        Path2D path = new Path2D.Double();
        if (shapeType == ShapeTypes.CHART_STAR) {
            path.moveTo(l, t);
            path.lineTo(r, b);
            path.moveTo(l, b);
            path.lineTo(r, t);
            path.moveTo(cx, t);
            path.lineTo(cx, b);
        } else if (shapeType == ShapeTypes.CHART_X) {
            path.moveTo(l, t);
            path.lineTo(r, b);
            path.moveTo(l, b);
            path.lineTo(r, t);
        } else {
            path.moveTo(cx, t);
            path.lineTo(cx, b);
            path.moveTo(l, cy);
            path.lineTo(r, cy);
        }
        return path;
    }

    private Shape cubeDetail(Rect rect, XSSFShape shape) {
        double ss = Math.min(rect.width, rect.height);
        double a = clamp(shapeAdjust(shape, "adj", 25000), 0, 100000);
        double y1 = ss * a / 100000d;
        double x4 = rect.width - y1;
        Path2D path = new Path2D.Double();
        path.moveTo(rect.x, rect.y + y1);
        path.lineTo(rect.x + x4, rect.y + y1);
        path.moveTo(rect.x + x4, rect.y + y1);
        path.lineTo(rect.x + rect.width, rect.y);
        path.moveTo(rect.x + x4, rect.y + y1);
        path.lineTo(rect.x + x4, rect.y + rect.height);
        return path;
    }

    private List<PresetFill> cubeFills(Rect rect, XSSFShape shape) {
        double ss = Math.min(rect.width, rect.height);
        double a = clamp(shapeAdjust(shape, "adj", 25000), 0, 100000);
        double y1 = ss * a / 100000d;
        double y4 = rect.height - y1;
        double x4 = rect.width - y1;
        double l = rect.x;
        double t = rect.y;
        double r = rect.x + rect.width;
        double b = rect.y + rect.height;

        Path2D top = new Path2D.Double();
        top.moveTo(l, t + y1);
        top.lineTo(l + y1, t);
        top.lineTo(r, t);
        top.lineTo(l + x4, t + y1);
        top.closePath();

        Path2D side = new Path2D.Double();
        side.moveTo(l + x4, t + y1);
        side.lineTo(r, t);
        side.lineTo(r, t + y4);
        side.lineTo(l + x4, b);
        side.closePath();

        return List.of(
                new PresetFill(top, "lightenLess"),
                new PresetFill(side, "darkenLess"));
    }

    private List<PresetFill> canFills(Rect rect, XSSFShape shape) {
        double ss = Math.min(rect.width, rect.height);
        double maxAdj = 50000d * rect.height / ss;
        double a = clamp(shapeAdjust(shape, "adj", 25000), 0, maxAdj);
        double y1 = ss * a / 200000d;
        Path2D top = new Path2D.Double();
        top.append(new Ellipse2D.Double(
                rect.x, rect.y, rect.width, y1 * 2d), false);
        return List.of(new PresetFill(top, "lighten"));
    }



    private Shape ribbonDetail(Rect rect, boolean inverted) {
        Path2D path = new Path2D.Double();
        addFoldLine(path, rect, 0.14, 0.38, 0.88);
        addFoldLine(path, rect, 0.86, 0.38, 0.88);
        return inverted ? flipVertical(path, rect.y + rect.height / 2d) : path;
    }


    private Shape ellipseRibbonDetail(Rect rect, boolean inverted) {
        Path2D path = new Path2D.Double();
        addFoldLine(path, rect, 0.08, 0.63, 0.90);
        addFoldLine(path, rect, 0.92, 0.63, 0.90);
        path.moveTo(rect.x + rect.width * 0.12, rect.y + rect.height * 0.57);
        path.quadTo(rect.x + rect.width / 2d, rect.y + rect.height * 0.12,
                rect.x + rect.width * 0.88, rect.y + rect.height * 0.57);
        return inverted ? flipVertical(path, rect.y + rect.height / 2d) : path;
    }


    private Shape leftRightRibbonDetail(Rect rect) {
        Path2D path = new Path2D.Double();
        addFoldLine(path, rect, 0.18, 0.25, 0.75);
        addFoldLine(path, rect, 0.82, 0.25, 0.75);
        return path;
    }


    private Shape scrollDetail(Rect rect, boolean vertical) {
        Path2D path = new Path2D.Double();
        if (vertical) {
            path.moveTo(rect.x + rect.width * 0.08, rect.y + rect.height * 0.20);
            path.lineTo(rect.x + rect.width * 0.92, rect.y + rect.height * 0.20);
            path.moveTo(rect.x + rect.width * 0.08, rect.y + rect.height * 0.80);
            path.lineTo(rect.x + rect.width * 0.92, rect.y + rect.height * 0.80);
            path.moveTo(rect.x + rect.width * 0.18, rect.y + rect.height * 0.20);
            path.lineTo(rect.x + rect.width * 0.18, rect.y + rect.height * 0.80);
            path.moveTo(rect.x + rect.width * 0.82, rect.y + rect.height * 0.20);
            path.lineTo(rect.x + rect.width * 0.82, rect.y + rect.height * 0.80);
        } else {
            path.moveTo(rect.x + rect.width * 0.20, rect.y + rect.height * 0.08);
            path.lineTo(rect.x + rect.width * 0.20, rect.y + rect.height * 0.92);
            path.moveTo(rect.x + rect.width * 0.80, rect.y + rect.height * 0.08);
            path.lineTo(rect.x + rect.width * 0.80, rect.y + rect.height * 0.92);
            path.moveTo(rect.x + rect.width * 0.20, rect.y + rect.height * 0.18);
            path.lineTo(rect.x + rect.width * 0.80, rect.y + rect.height * 0.18);
            path.moveTo(rect.x + rect.width * 0.20, rect.y + rect.height * 0.82);
            path.lineTo(rect.x + rect.width * 0.80, rect.y + rect.height * 0.82);
        }
        return path;
    }

    private static void addFoldLine(Path2D path, Rect rect, double x,
                                    double fromY, double toY) {
        path.moveTo(rect.x + rect.width * x, rect.y + rect.height * fromY);
        path.lineTo(rect.x + rect.width * x, rect.y + rect.height * toY);
    }


    private Shape foldedCornerFlap(Rect rect, double adjust) {
        double a = clamp(adjust, 0, 50000);
        double dy2 = Math.min(rect.width, rect.height) * a / 100000.0;
        double dy1 = dy2 / 5.0;
        double x1 = rect.x + rect.width - dy2;
        double y2 = rect.y + rect.height - dy2;
        Path2D path = new Path2D.Double();
        path.moveTo(x1, rect.y + rect.height);
        path.lineTo(x1 + dy1, y2 + dy1);
        path.lineTo(rect.x + rect.width, y2);
        path.closePath();
        return path;
    }

    /**
     * Preset outline: the flap triangle plus the main boundary. The cut edge
     * must only be stroked once; stroking the fill polygon and the flap
     * separately makes the fold look too heavy.
     */
    private Shape foldedCornerOutline(Rect rect, double adjust) {
        double a = clamp(adjust, 0, 50000);
        double dy2 = Math.min(rect.width, rect.height) * a / 100000.0;
        double dy1 = dy2 / 5.0;
        double x1 = rect.x + rect.width - dy2;
        double y2 = rect.y + rect.height - dy2;
        Path2D path = new Path2D.Double();
        path.moveTo(x1, rect.y + rect.height);
        path.lineTo(x1 + dy1, y2 + dy1);
        path.lineTo(rect.x + rect.width, y2);
        path.closePath();
        path.moveTo(x1, rect.y + rect.height);
        path.lineTo(rect.x, rect.y + rect.height);
        path.lineTo(rect.x, rect.y);
        path.lineTo(rect.x + rect.width, rect.y);
        path.lineTo(rect.x + rect.width, y2);
        return path;
    }

    private Shape borderCalloutLeader(XSSFSimpleShape shape, Rect rect, int segments) {
        double w = rect.width;
        double h = rect.height;
        double x1 = w * adjustValue(shape, "adj2", -8333) / 100000d;
        double y1 = h * adjustValue(shape, "adj1", 18750) / 100000d;
        double x2 = w * adjustValue(shape, "adj4", -38333) / 100000d;
        double y2 = h * adjustValue(shape, "adj3", 112500) / 100000d;
        Path2D path = new Path2D.Double();
        path.moveTo(rect.x + x1, rect.y + y1);
        path.lineTo(rect.x + x2, rect.y + y2);
        if (segments >= 2) {
            double x3 = w * adjustValue(shape, "adj6", -16667) / 100000d;
            double y3 = h * adjustValue(shape, "adj5", 18750) / 100000d;
            path.lineTo(rect.x + x3, rect.y + y3);
            if (segments >= 3) {
                double x4 = w * adjustValue(shape, "adj8", -8333) / 100000d;
                double y4 = h * adjustValue(shape, "adj7", 112963) / 100000d;
                path.lineTo(rect.x + x4, rect.y + y4);
            }
        }
        return path;
    }

    private Shape accentCalloutBar(Rect rect) {
        Path2D path = new Path2D.Double();
        path.moveTo(rect.x + rect.width * 0.10, rect.y + rect.height * 0.14);
        path.lineTo(rect.x + rect.width * 0.10, rect.y + rect.height * 0.86);
        path.moveTo(rect.x + rect.width * 0.16, rect.y + rect.height * 0.14);
        path.lineTo(rect.x + rect.width * 0.16, rect.y + rect.height * 0.86);
        return path;
    }



    private HorizontalAlignment toHorizontal(TextAlign align) {
        if (align == null) {
            return HorizontalAlignment.LEFT;
        }
        return switch (align) {
            case CENTER -> HorizontalAlignment.CENTER;
            case RIGHT -> HorizontalAlignment.RIGHT;
            default -> HorizontalAlignment.LEFT;
        };
    }

    private static CTTransform2D childTransform(XSSFShape child) {
        if (child instanceof XSSFSimpleShape simple) {
            return simple.getCTShape().getSpPr().getXfrm();
        }
        if (child instanceof XSSFConnector connector) {
            return connector.getCTConnector().getSpPr().getXfrm();
        }
        if (child instanceof XSSFPicture picture) {
            return picture.getCTPicture().getSpPr().getXfrm();
        }
        return null;
    }

    private static CTGroupTransform2D childGroupTransform(XSSFShape child) {
        if (child instanceof XSSFShapeGroup group) {
            return group.getCTGroupShape().getGrpSpPr().getXfrm();
        }
        return null;
    }

    private static double coordinate(Object value) {
        if (value instanceof Number number) {
            return number.doubleValue();
        }
        try {
            return Double.parseDouble(String.valueOf(value));
        } catch (RuntimeException e) {
            return 0;
        }
    }

    private record GroupTransform(double offsetX, double offsetY, double scaleX, double scaleY) {
        Rect rect(XSSFChildAnchor anchor) {
            return rect(anchor.getDx1(), anchor.getDy1(), anchor.getDx2(), anchor.getDy2());
        }

        Rect rect(double x1, double y1, double x2, double y2) {
            double sx1 = offsetX + x1 * scaleX;
            double sy1 = offsetY + y1 * scaleY;
            double sx2 = offsetX + x2 * scaleX;
            double sy2 = offsetY + y2 * scaleY;
            return new Rect(Math.min(sx1, sx2), Math.min(sy1, sy2),
                    Math.abs(sx2 - sx1), Math.abs(sy2 - sy1));
        }
    }

    private record Rect(double x, double y, double width, double height) {
    }

    private record ConnectorEndpoints(Point2D start, Point2D end, int startSite, int endSite) {
    }

    private record ConnectorRoute(Path2D path, List<Point2D> points) {
    }


    private record PresetFill(Shape shape, String mode) {
    }

}
