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
import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingTheme;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPrGeneral;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBrType;

import java.awt.Color;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Resolves WordprocessingML styles, numbering, tables and drawings into the
 * neutral block model used by the layout engine.
 */
final class DocxStyleResolver {

    private final XWPFDocument document;
    private final XWPFStyles styles;
    private final Map<String, XWPFStyle> styleMap = new HashMap<>();
    private final DocxNumbering numbering;
    private final DrawingTheme theme;
    private final FontResolver fonts;

    DocxStyleResolver(XWPFDocument document) {
        this.document = document;
        this.styles = document.getStyles();
        if (styles != null && styles.getStyles() != null) {
            for (XWPFStyle style : styles.getStyles()) {
                if (style.getStyleId() != null) {
                    styleMap.put(style.getStyleId(), style);
                }
            }
        }
        this.numbering = new DocxNumbering(document.getNumbering());
        this.theme = DocxThemeReader.read(document);
        this.fonts = new FontResolver();
    }

    DocxDocumentModel resolve() {
        List<DocxBlock> blocks = new ArrayList<>();
        for (IBodyElement element : document.getBodyElements()) {
            if (element instanceof XWPFParagraph paragraph) {
                blocks.add(paragraph(paragraph, false));
            } else if (element instanceof XWPFTable table) {
                blocks.add(table(table));
            }
        }
        return new DocxDocumentModel(
                DocxPageSetup.from(document.getDocument().getBody().getSectPr()),
                blocks, document.getHeaderFooterPolicy());
    }

    private DocxParagraphModel paragraph(XWPFParagraph paragraph, boolean tableCell) {
        DocxStyle style = baseStyle();
        String styleId = paragraph.getStyle();
        if (styleId == null || styleId.isBlank()) {
            styleId = "a";
        }
        applyStyleChain(style, styleId, new HashSet<>());
        style.applyParagraph(paragraph.getCTP().getPPr());

        List<DocxRunModel> runs = new ArrayList<>();
        if (!tableCell && style.numberingId != null) {
            DocxNumbering.NumberedText number = numbering.next(
                    style.numberingId, style.numberingLevel == null ? 0 : style.numberingLevel);
            if (number != null) {
                DocxStyle numberStyle = style.copy();
                numberStyle.leftIndent = Math.max(0, number.leftIndent());
                numberStyle.firstLineIndent = -Math.abs(number.hangingIndent());
                runs.add(new DocxRunModel(number.text() + " ", numberStyle, false, List.of()));
            }
        }

        for (XWPFRun run : paragraph.getRuns()) {
            DocxStyle runStyle = style.copy();
            CTRPr properties = run.getCTR().getRPr();
            if (properties != null && !properties.getRStyleList().isEmpty()) {
                String characterStyle = properties.getRStyleList().get(0).getVal();
                applyStyleChain(runStyle, characterStyle, new HashSet<>());
            }
            runStyle.applyRun(properties);
            boolean pageBreak = pageBreak(run);
            List<DocxPictureModel> pictures = new ArrayList<>();
            for (XWPFPicture picture : run.getEmbeddedPictures()) {
                pictures.add(new DocxPictureModel(picture.getPictureData(),
                        picture.getWidth(), picture.getDepth()));
            }
            String text = pageBreak ? "" : run.text();
            if (text != null && !text.isEmpty() || pageBreak || !pictures.isEmpty()) {
                runs.add(new DocxRunModel(text == null ? "" : text, runStyle,
                        pageBreak, List.copyOf(pictures)));
            }
        }

        List<DocxShape> shapes = new ArrayList<>();
        for (org.w3c.dom.Element drawing : DocxShapeAdapter.drawings(paragraph)) {
            if (hasCanvas(drawing)) {
                // Word canvas groups contain their own coordinate system and
                // multiple child shapes. They are outside the MVP shape range.
                continue;
            }
            try {
                DocxShape shape = DocxShapeAdapter.fromDrawing(drawing, theme, fonts, 1.0);
                if (shape != null) {
                    shapes.add(shape);
                }
            } catch (RuntimeException ignored) {
                // Unsupported drawing: keep the paragraph flowing.
            }
        }
        if (tableCell) {
            style.leftIndent = 0;
            style.rightIndent = 0;
            style.firstLineIndent = 0;
        }
        return new DocxParagraphModel(style, runs, style.pageBreakBefore, List.copyOf(shapes));
    }

    private DocxTableModel table(XWPFTable table) {
        CTTbl ct = table.getCTTbl();
        List<Double> widths = new ArrayList<>();
        if (ct.getTblGrid() != null) {
            CTTblGrid grid = ct.getTblGrid();
            for (CTTblGridCol column : grid.getGridColList()) {
                widths.add(number(column.getW()) / 20.0);
            }
        }
        if (widths.isEmpty() && !table.getRows().isEmpty()
                && !table.getRows().get(0).getTableCells().isEmpty()) {
            for (XWPFTableCell cell : table.getRows().get(0).getTableCells()) {
                double width = cell.getCTTc().getTcPr() == null
                        || cell.getCTTc().getTcPr().getTcW() == null
                        ? 120 : number(cell.getCTTc().getTcPr().getTcW().getW()) / 20.0;
                widths.add(width > 0 ? width : 120);
            }
        }
        double indent = ct.getTblPr() == null || ct.getTblPr().getTblInd() == null
                ? 0 : number(ct.getTblPr().getTblInd().getW()) / 20.0;
        Color borderColor = new Color(0xE6, 0xE6, 0xE6);
        double borderWidth = 0.5;
        if (!table.getRows().isEmpty() && !table.getRows().get(0).getTableCells().isEmpty()) {
            CTTcPr tcPr = table.getRows().get(0).getTableCells().get(0).getCTTc().getTcPr();
            if (tcPr != null && tcPr.getTcBorders() != null) {
                CTTcBorders borders = tcPr.getTcBorders();
                CTBorder border = borders.getTop() != null ? borders.getTop()
                        : borders.getLeft() != null ? borders.getLeft() : borders.getBottom();
                if (border != null) {
                    Color color = color(border);
                    if (color != null) {
                        borderColor = color;
                    }
                    if (border.getSz() != null) {
                        borderWidth = Math.max(0.25, border.getSz().intValue() / 8.0);
                    }
                }
            }
        }

        List<DocxTableModel.Row> rows = new ArrayList<>();
        for (XWPFTableRow row : table.getRows()) {
            List<DocxTableModel.Cell> cells = new ArrayList<>();
            for (XWPFTableCell cell : row.getTableCells()) {
                List<DocxParagraphModel> paragraphs = new ArrayList<>();
                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                    paragraphs.add(paragraph(paragraph, true));
                }
                if (paragraphs.isEmpty()) {
                    paragraphs.add(new DocxParagraphModel(new DocxStyle(), List.of(), false, List.of()));
                }
                cells.add(new DocxTableModel.Cell(List.copyOf(paragraphs)));
            }
            rows.add(new DocxTableModel.Row(List.copyOf(cells)));
        }
        return new DocxTableModel(indent, List.copyOf(widths), List.copyOf(rows),
                borderColor, borderWidth, 5.76, 1.0);
    }

    private DocxStyle baseStyle() {
        DocxStyle style = new DocxStyle();
        if (styles != null) {
            if (styles.getDefaultRunStyle() != null) {
                style.applyRun(styles.getDefaultRunStyle().getRPr());
            }
            if (styles.getDefaultParagraphStyle() != null) {
                style.applyParagraph(styles.getDefaultParagraphStyle().getPPr());
            }
        }
        return style;
    }

    private void applyStyleChain(DocxStyle style, String styleId, Set<String> visited) {
        if (styleId == null || styleId.isBlank() || !visited.add(styleId)) {
            return;
        }
        XWPFStyle wordStyle = styleMap.get(styleId);
        if (wordStyle == null) {
            return;
        }
        applyStyleChain(style, wordStyle.getBasisStyleID(), visited);
        CTStyle ct = wordStyle.getCTStyle();
        if (ct == null) {
            return;
        }
        if (ct.getRPr() != null) {
            style.applyRun(ct.getRPr());
        }
        if (ct.getPPr() != null) {
            style.applyParagraph(ct.getPPr());
        }
    }

    private static boolean hasCanvas(org.w3c.dom.Element drawing) {
        return drawing.getElementsByTagNameNS(
                "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
                "wpc").getLength() > 0;
    }

    private static boolean pageBreak(XWPFRun run) {
        if (run.getCTR() == null || run.getCTR().getBrList().isEmpty()) {
            return false;
        }
        for (CTBr br : run.getCTR().getBrList()) {
            if (br.getType() == STBrType.PAGE) {
                return true;
            }
        }
        return false;
    }

    private static Color color(CTBorder border) {
        if (border.getColor() == null) {
            return null;
        }
        String value = border.getColor().toString();
        if ("auto".equalsIgnoreCase(value) || value.length() != 6) {
            return Color.BLACK;
        }
        try {
            return new Color(Integer.parseInt(value, 16));
        } catch (NumberFormatException ignored) {
            return null;
        }
    }

    private static double number(Object value) {
        if (value == null) {
            return 0;
        }
        if (value instanceof Number number) {
            return number.doubleValue();
        }
        try {
            return new BigDecimal(value.toString()).doubleValue();
        } catch (NumberFormatException ignored) {
            return 0;
        }
    }
}
