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

import com.github.naofum.xlsx2pdf4j.internal.drawingml.CustomGeometry;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.CustomGeometryParser;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingTheme;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.PresetGeometry;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.PresetShape;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.PresetTextInsets;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.ShapeAdjustments;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.junit.jupiter.api.Test;

import javax.xml.parsers.DocumentBuilderFactory;
import java.awt.Shape;
import java.awt.geom.Area;
import java.io.ByteArrayInputStream;
import java.lang.reflect.Field;
import java.nio.charset.StandardCharsets;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * The DrawingML core must not depend on a host document. These tests exercise
 * it directly and verify the spreadsheet adapter maps every shape name.
 */
class DrawingMlCoreTest {

    @Test
    void everyPresetBuildsGeometryWithoutAHostDocument() {
        for (PresetShape preset : PresetShape.values()) {
            Shape shape = PresetGeometry.create(
                    preset, 10, 20, 120, 80, ShapeAdjustments.NONE);
            assertThat(shape)
                    .as("%s geometry", preset)
                    .isNotNull();
            assertThat(shape.getBounds2D().getWidth())
                    .as("%s width", preset)
                    .isGreaterThan(0);
            assertThat(shape.getBounds2D().getHeight())
                    .as("%s height", preset)
                    .isGreaterThan(0);
        }
    }

    @Test
    void customGeometryParserIsHostNeutral() throws Exception {
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

        CustomGeometry geometry = CustomGeometryParser.parse(
                document.getDocumentElement(), 10, 20, 100, 50);

        assertThat(geometry).isNotNull();
        assertThat(geometry.paths()).hasSize(1);
        assertThat(geometry.paths().get(0).fill()).isTrue();
        assertThat(geometry.paths().get(0).stroke()).isTrue();
        assertThat(geometry.paths().get(0).path().getBounds2D().getX())
                .isEqualTo(10, org.assertj.core.data.Offset.offset(0.001));
        assertThat(geometry.paths().get(0).path().getBounds2D().getWidth())
                .isEqualTo(100, org.assertj.core.data.Offset.offset(0.001));
        assertThat(geometry.paths().get(0).path().getBounds2D().getHeight())
                .isEqualTo(50, org.assertj.core.data.Offset.offset(0.001));
    }

    @Test
    void themeParsingIsHostNeutral() throws Exception {
        String xml = """
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <a:themeElements>
                    <a:clrScheme name="Test">
                      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
                      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
                      <a:accent1><a:srgbClr val="123456"/></a:accent1>
                    </a:clrScheme>
                  </a:themeElements>
                </a:theme>
                """;

        DrawingTheme theme = DrawingTheme.fromThemeXml(
                new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)));

        assertThat(theme.color("accent1")).isEqualTo(new java.awt.Color(0x12, 0x34, 0x56));
        assertThat(theme.color("bg1")).isEqualTo(java.awt.Color.WHITE);
        assertThat(theme.color("tx1")).isEqualTo(java.awt.Color.BLACK);
    }

    @Test
    void presetTextInsetsAreHostNeutral() {
        var rectangle = PresetTextInsets.forShape(
                PresetShape.RECT, 100, 100, ShapeAdjustments.NONE);
        assertThat(rectangle).isEqualTo(com.github.naofum.xlsx2pdf4j.internal.drawingml.TextInsets.NONE);

        var round = PresetTextInsets.forShape(
                PresetShape.ROUND_RECT, 100, 100, ShapeAdjustments.NONE);
        assertThat(round.left()).isGreaterThan(0);
        assertThat(round.top()).isEqualTo(round.left(), org.assertj.core.data.Offset.offset(0.001));

        var decision = PresetTextInsets.forShape(
                PresetShape.FLOW_CHART_DECISION, 100, 100, ShapeAdjustments.NONE);
        assertThat(decision).isEqualTo(
                new com.github.naofum.xlsx2pdf4j.internal.drawingml.TextInsets(25, 25, 25, 25));
    }

    @Test
    void adjustmentValuesChangePresetGeometry() {
        Shape plain = PresetGeometry.create(
                PresetShape.PLUS, 0, 0, 100, 100, ShapeAdjustments.NONE);
        Shape adjusted = PresetGeometry.create(
                PresetShape.PLUS, 0, 0, 100, 100,
                (name, fallback) -> "adj".equals(name) ? 40000 : fallback);

        assertThat(new Area(adjusted)).isNotEqualTo(new Area(plain));
    }

    @Test
    void spreadsheetShapeIdsMapToPresetNames() throws Exception {
        for (Field field : ShapeTypes.class.getFields()) {
            if (field.getType() != int.class) {
                continue;
            }
            try {
                PresetShape expected = PresetShape.valueOf(field.getName());
                assertThat(XlsxPresetShapes.from(field.getInt(null)))
                        .as(field.getName())
                        .isEqualTo(expected);
            } catch (IllegalArgumentException ignored) {
                // Constants without a corresponding preset are not supported.
            }
        }
    }
}
