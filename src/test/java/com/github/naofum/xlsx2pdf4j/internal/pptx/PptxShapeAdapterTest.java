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

package com.github.naofum.xlsx2pdf4j.internal.pptx;

import com.github.naofum.xlsx2pdf4j.internal.FontResolver;
import com.github.naofum.xlsx2pdf4j.internal.TextLayoutEngine;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingShape;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingStyle;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingTheme;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.PresetGeometry;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.PresetShape;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.ShapePainter;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.ShapeTextBody;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.presentationml.x2006.main.CTShape;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;

import static org.assertj.core.api.Assertions.assertThat;

class PptxShapeAdapterTest {

    @Test
    void adaptsGeometryStyleAndTextFromPresentationShape() {
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlide slide = ppt.createSlide();
            XSLFAutoShape shape = slide.createAutoShape();
            shape.setShapeType(ShapeType.ROUND_RECT);
            shape.setAnchor(new Rectangle2D.Double(50, 40, 200, 100));
            shape.setFillColor(new Color(0x33, 0x66, 0x99));
            shape.setLineColor(new Color(0x11, 0x22, 0x33));
            shape.setLineWidth(2.0);
            shape.setText("Hello");
            shape.setWordWrap(true);
            shape.setVerticalAlignment(org.apache.poi.sl.usermodel.VerticalAlignment.MIDDLE);
            XSLFTextRun run = shape.getTextParagraphs().get(0).getTextRuns().get(0);
            run.setFontSize(24.0);
            run.setBold(true);

            var guides = ((CTShape) shape.getXmlObject()).getSpPr().getPrstGeom().getAvLst();
            var guide = guides.addNewGd();
            guide.setName("adj");
            guide.setFmla("val 30000");

            DrawingTheme theme = PptxThemeReader.read(slide);
            assertThat(PptxPresetShapes.from(ShapeType.ROUND_RECT))
                    .isEqualTo(PresetShape.ROUND_RECT);
            assertThat(PptxShapeAdapter.presetShape(shape)).isEqualTo(PresetShape.ROUND_RECT);
            assertThat(PptxShapeAdapter.adjustments(shape).value("adj", 16667))
                    .isEqualTo(30000);

            DrawingStyle style = PptxShapeAdapter.style(shape, theme);
            assertThat(style.fill).isEqualTo(new Color(0x33, 0x66, 0x99));
            assertThat(style.line).isEqualTo(new Color(0x11, 0x22, 0x33));
            assertThat(style.lineWidth).isEqualTo(2.0);

            Rectangle2D frame = PptxShapeAdapter.frame(shape, 1.0);
            assertThat(frame.getWidth()).isEqualTo(200);
            assertThat(frame.getHeight()).isEqualTo(100);

            ShapeTextBody body = PptxShapeAdapter.textBody(
                    shape, theme, new FontResolver(), 1.0);
            assertThat(body).isNotNull();
            assertThat(body.text()).isEqualTo("Hello");
            assertThat(body.runs()).hasSize(1);
            assertThat(body.runs().get(0).font().getSize2D())
                    .isCloseTo(24.0f, org.assertj.core.data.Offset.offset(0.01f));
            assertThat(body.runs().get(0).font().isBold()).isTrue();
            assertThat(body.wrap()).isTrue();
            assertThat(body.vertical())
                    .isEqualTo(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

            var geometry = PresetGeometry.create(PresetShape.ROUND_RECT,
                    frame.getX(), frame.getY(), frame.getWidth(), frame.getHeight(),
                    PptxShapeAdapter.adjustments(shape));
            assertThat(geometry).isNotNull();
            assertThat(geometry.getBounds2D().getWidth()).isEqualTo(200);

            DrawingShape drawing = PptxShapeAdapter.drawingShape(
                    shape, theme, new FontResolver(), 1.0);
            assertThat(drawing).isNotNull();
            assertThat(drawing.preset()).isEqualTo(PresetShape.ROUND_RECT);
            assertThat(drawing.frame().getWidth()).isEqualTo(200);

            BufferedImage image = new BufferedImage(320, 180, BufferedImage.TYPE_INT_RGB);
            Graphics2D g2 = image.createGraphics();
            g2.setColor(Color.WHITE);
            g2.fillRect(0, 0, image.getWidth(), image.getHeight());
            assertThat(ShapePainter.draw(g2, drawing, new TextLayoutEngine(), null)).isTrue();
            g2.dispose();
            assertThat(nonWhitePixels(image)).isGreaterThan(0);
            assertThat(image.getRGB(80, 55) & 0xFFFFFF)
                    .isEqualTo(0x336699);
        } catch (Exception e) {
            throw new AssertionError(e);
        }
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
