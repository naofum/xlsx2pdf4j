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

import com.github.naofum.xlsx2pdf4j.api.ConversionWarning;
import com.github.naofum.xlsx2pdf4j.api.DocxConversionOptions;
import com.github.naofum.xlsx2pdf4j.api.ErrorCode;
import com.github.naofum.xlsx2pdf4j.api.TextLayerMode;
import com.github.naofum.xlsx2pdf4j.api.WarningCode;
import com.github.naofum.xlsx2pdf4j.api.Xlsx2PdfException;
import com.github.naofum.xlsx2pdf4j.internal.PdfWriter;
import com.github.naofum.xlsx2pdf4j.internal.PlacedText;
import com.github.naofum.xlsx2pdf4j.internal.TextLayoutEngine;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.ShapePainter;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFFieldRun;
import org.apache.poi.xwpf.usermodel.XWPFHeaderFooter;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.awt.Color;
import java.awt.Font;
import java.awt.FontMetrics;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.awt.geom.Rectangle2D;
import java.util.ArrayList;
import java.util.List;

/**
 * Flow layout for Word documents: paragraphs and tables are advanced from the
 * top of the content area and a new page is started when the cursor reaches
 * the bottom margin.
 */
final class DocxFlowLayout {

    private final DocxDocumentModel model;
    private final DocxConversionOptions options;
    private final PdfWriter pdf;
    private final List<ConversionWarning> warnings;
    private final DocxTextLayout textLayout = new DocxTextLayout();
    private final TextLayoutEngine shapeText = new TextLayoutEngine();
    private Page current;
    private int pageCount;

    DocxFlowLayout(DocxDocumentModel model, DocxConversionOptions options,
                   PdfWriter pdf, List<ConversionWarning> warnings) {
        this.model = model;
        this.options = options;
        this.pdf = pdf;
        this.warnings = warnings;
    }

    int render() throws Xlsx2PdfException {
        newPage();
        for (DocxBlock block : model.blocks()) {
            if (block instanceof DocxParagraphModel paragraph) {
                drawParagraph(paragraph);
            } else if (block instanceof DocxTableModel table) {
                drawTable(table);
            }
        }
        finishPage();
        return pageCount;
    }

    private void drawParagraph(DocxParagraphModel paragraph) throws Xlsx2PdfException {
        if (paragraph.pageBreakBefore() && current.y > bodyTop() + 0.1) {
            newPage();
        }
        List<DocxTextLayout.Line> lines = textLayout.layout(paragraph,
                model.pageSetup().contentWidth());
        if (lines.isEmpty()) {
            return;
        }
        double before = Math.max(0, paragraph.style().spaceBefore);
        if (current.y > bodyTop() + 0.1
                && current.y + before + lines.get(0).height() > bodyBottom()) {
            newPage();
        }
        current.y += before;
        for (DocxTextLayout.Line line : lines) {
            if (current.y > bodyTop() + 0.1
                    && current.y + line.height() > bodyBottom()) {
                newPage();
            }
            textLayout.draw(current.g2, line,
                    model.pageSetup().leftMargin() + line.leftIndent(),
                    current.y, current.textLayer);
            current.y += line.height();
            if (line.breakAfter()) {
                newPage();
            }
        }
        for (DocxShape shape : paragraph.shapes()) {
            drawShape(shape, paragraph.style().leftIndent);
        }
        current.y += Math.max(0, paragraph.style().spaceAfter);
    }

    private void drawTable(DocxTableModel table) throws Xlsx2PdfException {
        if (table.columnWidths().isEmpty()) {
            return;
        }
        double total = table.columnWidths().stream().mapToDouble(Double::doubleValue).sum();
        double available = model.pageSetup().contentWidth() - Math.max(0, table.indent());
        double scale = total > available && total > 0 ? available / total : 1.0;
        double x = model.pageSetup().leftMargin() + Math.max(0, table.indent());
        if (x + total * scale > model.pageSetup().width() - model.pageSetup().rightMargin()) {
            x = model.pageSetup().leftMargin();
        }
        for (DocxTableModel.Row row : table.rows()) {
            List<List<DocxTextLayout.Line>> cellLines = new ArrayList<>();
            double rowHeight = 0;
            for (int i = 0; i < row.cells().size(); i++) {
                double width = (i < table.columnWidths().size()
                        ? table.columnWidths().get(i) : 100) * scale;
                List<DocxTextLayout.Line> lines = new ArrayList<>();
                for (DocxParagraphModel paragraph : row.cells().get(i).paragraphs()) {
                    lines.addAll(textLayout.layout(paragraph,
                            Math.max(10, width - table.paddingX() * 2)));
                }
                cellLines.add(lines);
                double height = table.paddingY() * 2;
                for (DocxTextLayout.Line line : lines) {
                    height += line.height();
                }
                rowHeight = Math.max(rowHeight, height);
            }
            rowHeight = Math.max(rowHeight, 14);
            if (current.y > bodyTop() + 0.1 && current.y + rowHeight > bodyBottom()) {
                newPage();
            }
            double cellX = x;
            for (int i = 0; i < row.cells().size(); i++) {
                double width = (i < table.columnWidths().size()
                        ? table.columnWidths().get(i) : 100) * scale;
                current.g2.setColor(table.borderColor());
                current.g2.setStroke(new java.awt.BasicStroke((float) table.borderWidth()));
                current.g2.draw(new Rectangle2D.Double(cellX, current.y, width, rowHeight));
                double textY = current.y + table.paddingY();
                for (DocxTextLayout.Line line : cellLines.get(i)) {
                    textLayout.draw(current.g2, line, cellX + table.paddingX()
                            + line.leftIndent(), textY, current.textLayer);
                    textY += line.height();
                }
                cellX += width;
            }
            current.y += rowHeight;
        }
    }

    private void drawShape(DocxShape shape, double indent) throws Xlsx2PdfException {
        Rectangle2D frame = shape.frame();
        if (frame == null || frame.getWidth() <= 0 || frame.getHeight() <= 0) {
            return;
        }
        double x;
        double y;
        if (shape.inline()) {
            x = model.pageSetup().leftMargin() + Math.max(0, indent);
            y = current.y;
            if (y > bodyTop() + 0.1 && y + frame.getHeight() > bodyBottom()) {
                newPage();
                y = current.y;
            }
        } else {
            x = model.pageSetup().leftMargin() + frame.getX();
            y = bodyTop() + frame.getY();
        }
        try {
            ShapePainter.draw(current.g2, DocxShapeAdapter.drawingShape(shape),
                    shapeText, current.textLayer);
            if (shape.inline()) {
                current.y = y + frame.getHeight() + 4;
            }
        } catch (RuntimeException e) {
            warnings.add(new ConversionWarning(WarningCode.SHAPE_UNSUPPORTED,
                    "Page " + current.pageNumber, shape.name(),
                    "failed to draw shape: " + e.getMessage()));
        }
    }

    private void newPage() throws Xlsx2PdfException {
        if (current != null) {
            finishPage();
        }
        if (pageCount >= options.maxPages()) {
            throw new Xlsx2PdfException(ErrorCode.LIMIT_EXCEEDED,
                    "page count exceeds limit " + options.maxPages());
        }
        pageCount++;
        current = new Page(pageCount);
    }

    private void finishPage() throws Xlsx2PdfException {
        if (current == null) {
            return;
        }
        drawFooter();
        current.g2.dispose();
        pdf.addPage(current.image, model.pageSetup().width(), model.pageSetup().height(),
                current.textLayer);
        current.image.flush();
        current = null;
    }

    private void drawFooter() {
        XWPFHeaderFooterPolicy policy = model.headerFooterPolicy();
        if (policy == null) {
            return;
        }
        XWPFHeaderFooter footer;
        try {
            footer = policy.getFooter(current.pageNumber);
        } catch (RuntimeException e) {
            return;
        }
        String text = footerText(footer, current.pageNumber);
        if (text.isBlank()) {
            return;
        }
        Font font = new Font(Font.SANS_SERIF, Font.PLAIN, 9);
        current.g2.setFont(font);
        current.g2.setColor(new Color(0x7D, 0x7D, 0x7F));
        FontMetrics metrics = current.g2.getFontMetrics(font);
        double baseline = model.pageSetup().height() - model.pageSetup().footerDistance();
        current.g2.drawString(text, (float) model.pageSetup().leftMargin(),
                (float) baseline);
        if (current.textLayer != null) {
            current.textLayer.add(new PlacedText(text, model.pageSetup().leftMargin(),
                    baseline, 9));
        }
    }

    private String footerText(XWPFHeaderFooter footer, int pageNumber) {
        if (footer == null) {
            return "";
        }
        StringBuilder builder = new StringBuilder();
        for (XWPFParagraph paragraph : footer.getParagraphs()) {
            StringBuilder line = new StringBuilder();
            for (XWPFRun run : paragraph.getRuns()) {
                if (run instanceof XWPFFieldRun field
                        && field.getFieldInstruction() != null
                        && field.getFieldInstruction().toUpperCase(java.util.Locale.ROOT).contains("PAGE")) {
                    line.append(pageNumber);
                } else if (run.text() != null) {
                    line.append(run.text());
                }
            }
            if (line.length() > 0) {
                if (builder.length() > 0) {
                    builder.append(' ');
                }
                builder.append(line);
            }
        }
        String text = builder.toString()
                .replace("\u0007", "")
                .replace("\r", "")
                .replace("\n", " ")
                .trim();
        if (text.isEmpty()) {
            text = footer.getText() == null ? "" : footer.getText()
                    .replace("\u0007", "").replace("\r", "").replace("\n", " ").trim();
            // The footer text box often contains a cached PAGE field result.
            text = text.replaceFirst("(\\d+)(\\s*)$", pageNumber + "$2");
        }
        return text;
    }

    private double bodyTop() {
        return Math.max(model.pageSetup().topMargin(), model.pageSetup().headerDistance());
    }

    private double bodyBottom() {
        return model.pageSetup().height() - model.pageSetup().bottomMargin();
    }

    private final class Page {
        private final int pageNumber;
        private final BufferedImage image;
        private final Graphics2D g2;
        private final List<PlacedText> textLayer;
        private double y;

        private Page(int pageNumber) {
            this.pageNumber = pageNumber;
            double scale = options.dpi() / 72.0;
            this.image = new BufferedImage(
                    Math.max(1, (int) Math.round(model.pageSetup().width() * scale)),
                    Math.max(1, (int) Math.round(model.pageSetup().height() * scale)),
                    BufferedImage.TYPE_INT_RGB);
            this.g2 = image.createGraphics();
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING,
                    RenderingHints.VALUE_ANTIALIAS_ON);
            g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING,
                    RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
            g2.setColor(Color.WHITE);
            g2.fillRect(0, 0, image.getWidth(), image.getHeight());
            g2.scale(scale, scale);
            this.textLayer = options.textLayerMode() == TextLayerMode.INVISIBLE
                    ? new ArrayList<>() : null;
            this.y = bodyTop();
            drawHeader();
        }

        private void drawHeader() {
            XWPFHeaderFooterPolicy policy = model.headerFooterPolicy();
            if (policy == null) {
                return;
            }
            XWPFHeaderFooter header;
            try {
                header = policy.getHeader(pageNumber);
            } catch (RuntimeException e) {
                return;
            }
            String text = headerText(header);
            if (text.isBlank()) {
                return;
            }
            Font font = new Font(Font.SANS_SERIF, Font.PLAIN, 9);
            g2.setFont(font);
            g2.setColor(new Color(0x7D, 0x7D, 0x7F));
            double baseline = model.pageSetup().headerDistance() + font.getSize2D();
            g2.drawString(text, (float) model.pageSetup().leftMargin(), (float) baseline);
            if (textLayer != null) {
                textLayer.add(new PlacedText(text, model.pageSetup().leftMargin(),
                        baseline, font.getSize2D()));
            }
            y = Math.max(y, baseline + 4);
        }

        private String headerText(XWPFHeaderFooter header) {
            if (header == null) {
                return "";
            }
            StringBuilder builder = new StringBuilder();
            for (XWPFParagraph paragraph : header.getParagraphs()) {
                String text = paragraph.getText();
                if (text != null && !text.isBlank()) {
                    if (builder.length() > 0) {
                        builder.append(' ');
                    }
                    builder.append(text.replace("\r", "").replace("\n", " "));
                }
            }
            return builder.toString().replace("\u0007", "").trim();
        }
    }
}
