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

import com.github.naofum.xlsx2pdf4j.internal.TextLayoutEngine;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.junit.jupiter.api.Test;

import javax.xml.parsers.DocumentBuilderFactory;
import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * The common painter must render a host-neutral shape and custom geometry
 * without any workbook, presentation or document object.
 */
class ShapePainterTest {

    @Test
    void paintsPresetFillLineAndText() {
        DrawingStyle style = new DrawingStyle();
        style.fill = new Color(0x33, 0x66, 0x99);
        style.line = new Color(0x11, 0x22, 0x33);
        style.lineWidth = 2;

        DrawingShape shape = DrawingShape.preset("rect",
                new java.awt.geom.Rectangle2D.Double(10, 10, 100, 60),
                PresetShape.RECT, ShapeAdjustments.NONE, style,
                textBody("Hi"), null);

        BufferedImage image = new BufferedImage(140, 100, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2 = image.createGraphics();
        g2.setColor(Color.WHITE);
        g2.fillRect(0, 0, image.getWidth(), image.getHeight());
        assertThat(ShapePainter.draw(g2, shape, new TextLayoutEngine(), null)).isTrue();
        g2.dispose();

        assertThat(image.getRGB(20, 20) & 0xFFFFFF).isEqualTo(0x336699);
        assertThat(nonWhitePixels(image)).isGreaterThan(500);
    }

    @Test
    void paintsCustomGeometryAndItsPerPathFill() throws Exception {
        String xml = """
                <a:custGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <a:pathLst>
                    <a:path w="100" h="100">
                      <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
                      <a:lnTo><a:pt x="100" y="0"/></a:lnTo>
                      <a:lnTo><a:pt x="100" y="100"/></a:lnTo>
                      <a:lnTo><a:pt x="0" y="100"/></a:lnTo>
                      <a:close/>
                    </a:path>
                  </a:pathLst>
                </a:custGeom>
                """;
        var factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        var document = factory.newDocumentBuilder().parse(
                new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)));
        CustomGeometry custom = CustomGeometryParser.parse(
                document.getDocumentElement(), 10, 10, 100, 50);

        DrawingStyle style = new DrawingStyle();
        style.fill = new Color(0x22, 0x88, 0x44);
        style.line = null;
        style.lineNone = true;

        DrawingShape shape = DrawingShape.custom("custom",
                new java.awt.geom.Rectangle2D.Double(10, 10, 100, 50),
                custom, style, null, null);

        BufferedImage image = new BufferedImage(140, 80, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2 = image.createGraphics();
        g2.setColor(Color.WHITE);
        g2.fillRect(0, 0, image.getWidth(), image.getHeight());
        assertThat(ShapePainter.draw(g2, shape, new TextLayoutEngine(), null)).isTrue();
        g2.dispose();

        assertThat(image.getRGB(60, 35) & 0xFFFFFF).isEqualTo(0x228844);
    }

    private static ShapeTextBody textBody(String text) {
        Font font = new Font(Font.SANS_SERIF, Font.PLAIN, 12);
        TextLayoutEngine.StyledRun run = new TextLayoutEngine.StyledRun(
                text, font, Color.BLACK, false, false);
        TextLayoutEngine.Paragraph paragraph = new TextLayoutEngine.Paragraph(
                List.of(run), HorizontalAlignment.CENTER, font);
        return new ShapeTextBody(List.of(paragraph), List.of(run), text,
                new TextInsets(2, 2, 2, 2), VerticalAlignment.CENTER,
                true, false, false, false, false, 1.0);
    }

    private static int nonWhitePixels(BufferedImage image) {
        int count = 0;
        for (int y = 0; y < image.getHeight(); y++) {
            for (int x = 0; x < image.getWidth(); x++) {
                if ((image.getRGB(x, y) & 0xFFFFFF) != 0xFFFFFF) {
                    count++;
                }
            }
        }
        return count;
    }
}
