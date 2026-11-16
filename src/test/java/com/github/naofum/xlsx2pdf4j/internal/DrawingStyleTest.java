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

import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingStyle;
import org.junit.jupiter.api.Test;

import java.awt.Color;

import static org.assertj.core.api.Assertions.assertThat;

class DrawingStyleTest {

    @Test
    void parsesGradientAndOuterShadow() {
        String xml = """
                <xml-fragment xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <a:gradFill>
                    <a:gsLst>
                      <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
                      <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
                    </a:gsLst>
                    <a:lin ang="5400000"/>
                  </a:gradFill>
                  <a:effectLst>
                    <a:outerShdw dist="12700" dir="0">
                      <a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr>
                    </a:outerShdw>
                  </a:effectLst>
                </xml-fragment>
                """;

        DrawingStyle style = DrawingStyle.parse(xml);

        assertThat(style.gradientStops).hasSize(2);
        assertThat(style.gradientAngle).isEqualTo(90);
        assertThat(style.shadow).isNotNull();
        assertThat(style.shadow.dx()).isEqualTo(1.0);
        assertThat(style.shadow.dy()).isEqualTo(0.0, org.assertj.core.data.Offset.offset(0.001));
        assertThat(style.shadow.color().getAlpha()).isBetween(120, 135);
        assertThat(style.shadow.color().getRed()).isEqualTo(0);
        assertThat(style.shadow.color() instanceof Color).isTrue();
    }
}
