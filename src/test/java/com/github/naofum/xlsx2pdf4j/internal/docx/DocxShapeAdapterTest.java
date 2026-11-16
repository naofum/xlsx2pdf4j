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
import com.github.naofum.xlsx2pdf4j.internal.drawingml.PresetGeometry;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.PresetShape;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.ShapePainter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.nio.charset.StandardCharsets;

import static org.assertj.core.api.Assertions.assertThat;

class DocxShapeAdapterTest {

    @Test
    void adaptsDrawingMlShapeFromWordParagraph() throws Exception {
        String xml = """
                <w:drawing
                    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                  <wp:inline>
                    <wp:extent cx="2540000" cy="1270000"/>
                    <wp:docPr id="1" name="Rounded"/>
                    <a:graphic>
                      <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                        <wps:wsp>
                          <wps:cNvSpPr/>
                          <wps:spPr>
                            <a:xfrm><a:off x="0" y="0"/><a:ext cx="2540000" cy="1270000"/></a:xfrm>
                            <a:prstGeom prst="roundRect">
                              <a:avLst><a:gd name="adj" fmla="val 30000"/></a:avLst>
                            </a:prstGeom>
                            <a:solidFill><a:srgbClr val="336699"/></a:solidFill>
                            <a:ln w="25400"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln>
                          </wps:spPr>
                          <wps:txbx>
                            <w:txbxContent>
                              <w:p>
                                <w:pPr><w:jc w:val="center"/></w:pPr>
                                <w:r>
                                  <w:rPr>
                                    <w:rFonts w:ascii="Arial"/>
                                    <w:sz w:val="48"/>
                                    <w:b/>
                                    <w:color w:val="FF0000"/>
                                  </w:rPr>
                                  <w:t>Hello</w:t>
                                </w:r>
                              </w:p>
                            </w:txbxContent>
                          </wps:txbx>
                          <wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720"
                                      anchor="ctr" wrap="square"/>
                        </wps:wsp>
                      </a:graphicData>
                    </a:graphic>
                  </wp:inline>
                </w:drawing>
                """;

        try (XWPFDocument document = new XWPFDocument()) {
            XWPFParagraph paragraph = document.createParagraph();
            CTDrawing drawing = CTDrawing.Factory.parse(xml);
            paragraph.getCTP().addNewR().setDrawingArray(new CTDrawing[]{drawing});

            var drawings = DocxShapeAdapter.drawings(paragraph);
            assertThat(drawings).hasSize(1);
            assertThat(DocxPresetShapes.fromOoxml("roundRect")).isEqualTo(PresetShape.ROUND_RECT);
            assertThat(DocxPresetShapes.fromOoxml("flowChartDecision"))
                    .isEqualTo(PresetShape.FLOW_CHART_DECISION);

            DocxShape shape = DocxShapeAdapter.fromDrawing(
                    drawings.get(0), DocxThemeReader.read(document), new FontResolver(), 1.0);
            assertThat(shape).isNotNull();
            assertThat(shape.name()).isEqualTo("Rounded");
            assertThat(shape.preset()).isEqualTo(PresetShape.ROUND_RECT);
            assertThat(shape.inline()).isTrue();
            assertThat(shape.adjustments().value("adj", 16667)).isEqualTo(30000);
            assertThat(shape.frame().getWidth()).isEqualTo(200);
            assertThat(shape.frame().getHeight()).isEqualTo(100);
            assertThat(shape.style().fill).isEqualTo(new Color(0x33, 0x66, 0x99));
            assertThat(shape.style().line).isEqualTo(new Color(0x11, 0x22, 0x33));
            assertThat(shape.style().lineWidth).isEqualTo(2.0);

            assertThat(shape.text()).isNotNull();
            assertThat(shape.text().text()).isEqualTo("Hello");
            assertThat(shape.text().runs()).hasSize(1);
            assertThat(shape.text().runs().get(0).font().getSize2D())
                    .isCloseTo(24.0f, org.assertj.core.data.Offset.offset(0.01f));
            assertThat(shape.text().runs().get(0).font().isBold()).isTrue();
            assertThat(shape.text().runs().get(0).color()).isEqualTo(Color.RED);
            assertThat(shape.text().vertical())
                    .isEqualTo(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

            var geometry = PresetGeometry.create(shape.preset(),
                    shape.frame().getX(), shape.frame().getY(),
                    shape.frame().getWidth(), shape.frame().getHeight(),
                    shape.adjustments());
            assertThat(geometry).isNotNull();
            assertThat(geometry.getBounds2D().getWidth()).isEqualTo(200);

            DrawingShape drawingShape = DocxShapeAdapter.drawingShape(shape);
            assertThat(drawingShape).isNotNull();
            assertThat(drawingShape.preset()).isEqualTo(PresetShape.ROUND_RECT);
            assertThat(drawingShape.frame().getWidth()).isEqualTo(200);

            BufferedImage image = new BufferedImage(320, 180, BufferedImage.TYPE_INT_RGB);
            Graphics2D g2 = image.createGraphics();
            g2.setColor(Color.WHITE);
            g2.fillRect(0, 0, image.getWidth(), image.getHeight());
            assertThat(ShapePainter.draw(g2, drawingShape, new TextLayoutEngine(), null)).isTrue();
            g2.dispose();
            assertThat(nonWhitePixels(image)).isGreaterThan(0);
            assertThat(image.getRGB(20, 15) & 0xFFFFFF)
                    .isEqualTo(0x336699);
        }
    }

    @Test
    void ooxmlPresetNamesBecomeEnumNames() {
        assertThat(DocxPresetShapes.enumName("roundRect")).isEqualTo("ROUND_RECT");
        assertThat(DocxPresetShapes.enumName("star5")).isEqualTo("STAR_5");
        assertThat(DocxPresetShapes.enumName("nonIsoscelesTrapezoid"))
                .isEqualTo("NON_ISOSCELES_TRAPEZOID");
        assertThat(new String("x".getBytes(StandardCharsets.UTF_8), StandardCharsets.UTF_8))
                .isEqualTo("x");
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
