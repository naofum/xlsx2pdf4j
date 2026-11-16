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

package com.github.naofum.xlsx2pdf4j.internal.chart;

import com.github.naofum.xlsx2pdf4j.internal.FontResolver;
import com.github.naofum.xlsx2pdf4j.internal.PlacedText;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingTheme;
import org.apache.poi.xddf.usermodel.chart.XDDFChart;

import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.util.List;

/**
 * Entry point for drawing a spreadsheet or presentation chart onto a page.
 */
public final class ChartRenderer {

    private final FontResolver fonts;

    public ChartRenderer(FontResolver fonts) {
        this.fonts = fonts;
    }

    public boolean draw(Graphics2D g2, XDDFChart chart, Rectangle2D frame,
                        DrawingTheme theme, List<PlacedText> textLayer) {
        ChartSpec spec = XlsxChartReader.read(chart, theme);
        if (spec == null) {
            return false;
        }
        return XChartPainter.paint(spec, g2, frame, fonts, theme);
    }
}
