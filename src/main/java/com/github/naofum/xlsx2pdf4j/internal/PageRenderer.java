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

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.util.List;

/**
 * Renders one page into a BufferedImage at the requested DPI.
 */
public final class PageRenderer {

    private final CellRenderer cells;
    private final ShapeRenderer shapes;
    private final HeaderFooterRenderer headerFooter;
    private final int dpi;

    public PageRenderer(CellRenderer cells, ShapeRenderer shapes, HeaderFooterRenderer headerFooter, int dpi) {
        this.cells = cells;
        this.shapes = shapes;
        this.headerFooter = headerFooter;
        this.dpi = dpi;
    }

    public BufferedImage render(PageLayout page, int totalPages, String fileName,
                                List<ConversionWarning> warnings, List<PlacedText> textLayer) {
        int width = Units.pointsToPixels(page.pageWidth, dpi);
        int height = Units.pointsToPixels(page.pageHeight, dpi);
        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2 = image.createGraphics();
        try {
            g2.setColor(Color.WHITE);
            g2.fillRect(0, 0, width, height);
            g2.scale(dpi / 72d, dpi / 72d);
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
            g2.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
            g2.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
            g2.clip(new java.awt.geom.Rectangle2D.Double(0, 0, page.pageWidth, page.pageHeight));

            drawTitles(g2, page, textLayer);
            cells.drawRange(g2, page.firstRow, page.lastRow, page.firstColumn, page.lastColumn,
                    page.contentLeft, page.contentTop, page.scale, textLayer);
            shapes.drawShapes(g2, page, warnings, textLayer);
            headerFooter.render(g2, page, totalPages, fileName, textLayer);
        } finally {
            g2.dispose();
        }
        return image;
    }

    private void drawTitles(Graphics2D g2, PageLayout page, List<PlacedText> textLayer) {
        SheetLayout sheet = page.sheet;
        PrintSettingsModel print = sheet.print;
        if (sheet.titleFirstRow >= 0) {
            double x = sheet.titleFirstColumn >= 0 ? print.marginLeft : page.contentLeft;
            double y = page.contentTop - sheet.renderTitleHeight() * page.scale;
            cells.drawRange(g2, sheet.titleFirstRow, sheet.titleLastRow,
                    sheet.firstColumn, sheet.lastColumn, x, y, page.scale, textLayer);
        }
        if (sheet.titleFirstColumn >= 0) {
            double x = print.marginLeft;
            double y = page.contentTop;
            cells.drawRange(g2, page.firstRow, page.lastRow,
                    sheet.titleFirstColumn, sheet.titleLastColumn, x, y, page.scale, textLayer);
        }
    }
}
