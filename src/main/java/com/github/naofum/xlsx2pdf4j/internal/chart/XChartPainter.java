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
import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingTheme;
import org.knowm.xchart.CategoryChart;
import org.knowm.xchart.CategoryChartBuilder;
import org.knowm.xchart.CategorySeries;
import org.knowm.xchart.HorizontalBarChart;
import org.knowm.xchart.HorizontalBarChartBuilder;
import org.knowm.xchart.HorizontalBarSeries;
import org.knowm.xchart.PieChart;
import org.knowm.xchart.PieChartBuilder;
import org.knowm.xchart.PieSeries;
import org.knowm.xchart.style.CategoryStyler;
import org.knowm.xchart.style.HorizontalBarStyler;
import org.knowm.xchart.style.PieStyler;
import org.knowm.xchart.style.Styler;
import org.knowm.xchart.style.lines.SeriesLines;
import org.knowm.xchart.style.markers.SeriesMarkers;

import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Paints the shared chart model with XChart onto the page Graphics2D.
 */
public final class XChartPainter {

    private XChartPainter() {
    }

    public static boolean paint(ChartSpec spec, Graphics2D g2, Rectangle2D frame,
                                FontResolver fonts, DrawingTheme theme) {
        if (spec == null || frame == null || frame.getWidth() <= 0 || frame.getHeight() <= 0) {
            return false;
        }
        if (spec.kind() == ChartKind.PIE || spec.kind() == ChartKind.DOUGHNUT) {
            return paintPie(spec, g2, frame, fonts, theme);
        }
        if (spec.kind() == ChartKind.BAR) {
            return paintHorizontalBar(spec, g2, frame, fonts, theme);
        }
        if (spec.kind() == ChartKind.SCATTER) {
            // Scatter needs a separate XChart chart type.
            return false;
        }
        return paintCategory(spec, g2, frame, fonts, theme);
    }

    private static boolean paintCategory(ChartSpec spec, Graphics2D g2, Rectangle2D frame,
                                         FontResolver fonts, DrawingTheme theme) {
        int width = Math.max(1, (int) Math.round(frame.getWidth()));
        int height = Math.max(1, (int) Math.round(frame.getHeight()));
        CategoryChart chart = new CategoryChartBuilder()
                .width(width)
                .height(height)
                .title(spec.title() == null ? "" : spec.title())
                .build();
        CategoryStyler styler = chart.getStyler();
        Font base = baseFont(fonts, spec);
        styler.setBaseFont(base);
        styler.setChartTitleFont(base.deriveFont(Font.BOLD, 12f));
        styler.setAxisTickLabelsFont(base.deriveFont(9f));
        styler.setLegendFont(base.deriveFont(9f));
        styler.setChartBackgroundColor(Color.WHITE);
        styler.setPlotBackgroundColor(Color.WHITE);
        styler.setLegendVisible(spec.legendVisible());
        styler.setLegendPosition(legendPosition(spec.legendPosition()));
        styler.setPlotGridLinesVisible(true);
        styler.setPlotGridVerticalLinesVisible(false);
        styler.setAvailableSpaceFill(0.8);
        styler.setOverlapped(spec.stacked());
        styler.setStacked(spec.stacked());

        ChartAxis primary = spec.valueAxes().isEmpty()
                ? new ChartAxis(null, null, null, null, null, false, true)
                : spec.valueAxes().get(0);
        applyYAxis(styler, 0, primary, spec.percentStacked());
        for (int i = 1; i < spec.valueAxes().size(); i++) {
            styler.setYAxisGroupPosition(i, Styler.YAxisPosition.Right);
            applyYAxis(styler, i, spec.valueAxes().get(i), spec.percentStacked());
        }

        List<ChartSeries> seriesList = spec.series();
        double[] totals = spec.percentStacked() ? totals(seriesList) : null;
        boolean reverse = spec.categoryReversed();
        for (ChartSeries series : seriesList) {
            double[] values = totals == null ? series.values()
                    : percentValues(series.values(), totals);
            CategorySeries xchartSeries = chart.addSeries(
                    series.name(), orderedCategories(series, reverse),
                    boxedValues(values, reverse, false));
            xchartSeries.setChartCategorySeriesRenderStyle(renderStyle(series.kind()));
            xchartSeries.setFillColor(series.fill());
            xchartSeries.setLineColor(series.line());
            xchartSeries.setLineWidth((float) Math.max(0.5, series.lineWidth()));
            xchartSeries.setMarker(series.kind() == ChartKind.LINE
                    ? SeriesMarkers.CIRCLE : SeriesMarkers.NONE);
            xchartSeries.setLineStyle(SeriesLines.SOLID);
            xchartSeries.setSmooth(series.smooth());
            xchartSeries.setYAxisGroup(series.axisGroup());
        }

        Graphics2D copy = (Graphics2D) g2.create();
        try {
            copy.translate(frame.getX(), frame.getY());
            copy.clip(new Rectangle2D.Double(0, 0, width, height));
            chart.paint(copy, width, height);
        } finally {
            copy.dispose();
        }
        return true;
    }

    private static boolean paintHorizontalBar(ChartSpec spec, Graphics2D g2, Rectangle2D frame,
                                              FontResolver fonts, DrawingTheme theme) {
        if (spec.series().isEmpty() || spec.stacked() || spec.percentStacked()
                || spec.valueAxes().size() > 1) {
            // XChart's horizontal bar chart has no stacked variant and no
            // secondary value axis.
            return false;
        }
        for (ChartSeries series : spec.series()) {
            if (series.kind() != ChartKind.BAR) {
                // Mixed horizontal bar/line charts are not supported.
                return false;
            }
        }
        int width = Math.max(1, (int) Math.round(frame.getWidth()));
        int height = Math.max(1, (int) Math.round(frame.getHeight()));
        HorizontalBarChart chart = new HorizontalBarChartBuilder()
                .width(width)
                .height(height)
                .title(spec.title() == null ? "" : spec.title())
                .build();
        HorizontalBarStyler styler = chart.getStyler();
        Font base = baseFont(fonts, spec);
        styler.setBaseFont(base);
        styler.setChartTitleFont(base.deriveFont(Font.BOLD, 12f));
        styler.setAxisTickLabelsFont(base.deriveFont(9f));
        styler.setLegendFont(base.deriveFont(9f));
        styler.setChartBackgroundColor(Color.WHITE);
        styler.setPlotBackgroundColor(Color.WHITE);
        styler.setLegendVisible(spec.legendVisible());
        styler.setLegendPosition(legendPosition(spec.legendPosition()));
        styler.setPlotGridLinesVisible(true);
        styler.setPlotGridVerticalLinesVisible(true);
        styler.setPlotGridHorizontalLinesVisible(false);
        styler.setAvailableSpaceFill(0.8);
        // XChart applies the X max-label cap to the vertical category axis of
        // a horizontal bar chart. When all category labels fit, disabling
        // label skipping keeps them aligned with the bars without changing
        // the numeric value-axis tick density.
        int categoryCount = maxCategoryCount(spec);
        int labelSpacing = Math.max(8,
                (int) Math.ceil(styler.getAxisTickLabelsFont().getSize2D() + 4));
        int labelsThatFit = Math.max(2, (int) Math.floor(height * 0.6 / labelSpacing));
        if (categoryCount <= labelsThatFit) {
            styler.setXAxisMaxLabelCount(Math.max(2, categoryCount));
        }
        ChartAxis primary = spec.valueAxes().isEmpty()
                ? new ChartAxis(null, null, null, null, null, false, true)
                : spec.valueAxes().get(0);
        applyXAxis(styler, primary);

        boolean reverse = !spec.categoryReversed();
        for (ChartSeries series : spec.series()) {
            HorizontalBarSeries xchartSeries = chart.addSeries(
                    series.name(), boxedValues(series.values(), reverse, true),
                    orderedCategories(series, reverse));
            xchartSeries.setFillColor(series.fill());
            xchartSeries.setLineColor(series.line());
            xchartSeries.setLineWidth((float) Math.max(0.5, series.lineWidth()));
        }

        Graphics2D copy = (Graphics2D) g2.create();
        try {
            copy.translate(frame.getX(), frame.getY());
            copy.clip(new Rectangle2D.Double(0, 0, width, height));
            chart.paint(copy, width, height);
        } finally {
            copy.dispose();
        }
        return true;
    }

    private static boolean paintPie(ChartSpec spec, Graphics2D g2, Rectangle2D frame,
                                    FontResolver fonts, DrawingTheme theme) {
        if (spec.series().isEmpty()) {
            return false;
        }
        int width = Math.max(1, (int) Math.round(frame.getWidth()));
        int height = Math.max(1, (int) Math.round(frame.getHeight()));
        PieChart chart = new PieChartBuilder()
                .width(width)
                .height(height)
                .title(spec.title() == null ? "" : spec.title())
                .build();
        PieStyler styler = chart.getStyler();
        Font base = baseFont(fonts, spec);
        styler.setBaseFont(base);
        styler.setChartTitleFont(base.deriveFont(Font.BOLD, 12f));
        styler.setLabelsFont(base.deriveFont(9f));
        styler.setLegendFont(base.deriveFont(9f));
        styler.setChartBackgroundColor(Color.WHITE);
        styler.setLegendVisible(spec.legendVisible());
        styler.setLegendPosition(legendPosition(spec.legendPosition()));
        styler.setCircular(true);
        styler.setLabelsVisible(false);
        styler.setStartAngleInDegrees(spec.firstSliceAngle());
        if (spec.kind() == ChartKind.DOUGHNUT) {
            styler.setDefaultSeriesRenderStyle(PieSeries.PieSeriesRenderStyle.Donut);
            double hole = spec.holeSize();
            if (hole > 1.0) {
                hole /= 100.0;
            }
            styler.setDonutThickness(Math.max(0.05, Math.min(0.95, hole)));
        }

        ChartSeries series = spec.series().get(0);
        for (int i = 0; i < Math.min(series.categories().size(), series.values().length); i++) {
            double value = series.values()[i];
            if (Double.isNaN(value)) {
                value = 0;
            }
            PieSeries slice = chart.addSeries(series.categories().get(i), value);
            slice.setFillColor(sliceColor(theme, i));
            slice.setChartPieSeriesRenderStyle(spec.kind() == ChartKind.DOUGHNUT
                    ? PieSeries.PieSeriesRenderStyle.Donut
                    : PieSeries.PieSeriesRenderStyle.Pie);
        }

        Graphics2D copy = (Graphics2D) g2.create();
        try {
            copy.translate(frame.getX(), frame.getY());
            copy.clip(new Rectangle2D.Double(0, 0, width, height));
            chart.paint(copy, width, height);
        } finally {
            copy.dispose();
        }
        return true;
    }

    private static void applyYAxis(CategoryStyler styler, int group,
                                   ChartAxis axis, boolean percentStacked) {
        if (axis == null) {
            return;
        }
        if (axis.min() != null) {
            styler.setYAxisMin(group, axis.min());
        }
        if (axis.max() != null) {
            styler.setYAxisMax(group, axis.max());
        }
        String pattern = percentStacked ? "0%" : decimalPattern(axis.numberFormat());
        if (pattern != null) {
            if (group == 0) {
                styler.setYAxisDecimalPattern(pattern);
            } else {
                styler.putYAxisGroupDecimalPatternMap(group, pattern);
            }
        }
        if (axis.logarithmic()) {
            styler.setYAxisLogarithmic(true);
        }
    }

    private static void applyXAxis(HorizontalBarStyler styler, ChartAxis axis) {
        if (axis == null) {
            return;
        }
        if (axis.min() != null) {
            styler.setXAxisMin(axis.min());
        }
        if (axis.max() != null) {
            styler.setXAxisMax(axis.max());
        }
        String pattern = decimalPattern(axis.numberFormat());
        if (pattern != null) {
            styler.setXAxisDecimalPattern(pattern);
        }
        if (axis.logarithmic()) {
            styler.setXAxisLogarithmic(true);
        }
    }

    private static int maxCategoryCount(ChartSpec spec) {
        int count = 0;
        for (ChartSeries series : spec.series()) {
            count = Math.max(count, series.categories().size());
        }
        return count;
    }

    private static CategorySeries.CategorySeriesRenderStyle renderStyle(ChartKind kind) {
        return switch (kind) {
            case LINE -> CategorySeries.CategorySeriesRenderStyle.Line;
            case AREA -> CategorySeries.CategorySeriesRenderStyle.Area;
            case SCATTER -> CategorySeries.CategorySeriesRenderStyle.Scatter;
            default -> CategorySeries.CategorySeriesRenderStyle.Bar;
        };
    }

    private static Styler.LegendPosition legendPosition(String position) {
        if (position == null) {
            return Styler.LegendPosition.OutsideE;
        }
        return switch (position) {
            case "bottom" -> Styler.LegendPosition.OutsideS;
            case "top" -> Styler.LegendPosition.InsideN;
            case "left" -> Styler.LegendPosition.InsideNW;
            default -> Styler.LegendPosition.OutsideE;
        };
    }

    private static Font baseFont(FontResolver fonts, ChartSpec spec) {
        StringBuilder sample = new StringBuilder(spec.title() == null ? "" : spec.title());
        for (ChartSeries series : spec.series()) {
            sample.append(' ').append(series.name());
            for (String category : series.categories()) {
                sample.append(' ').append(category);
            }
        }
        return fonts.resolveForText("Arial", 10, false, false, sample.toString());
    }

    private static String decimalPattern(String format) {
        if (format == null || format.isBlank()) {
            return null;
        }
        String pattern = format;
        int section = pattern.indexOf(';');
        if (section >= 0) {
            pattern = pattern.substring(0, section);
        }
        pattern = pattern.replaceAll("\\[[^]]*]", "")
                .replace("\"", "")
                .replace("_", "")
                .replace("*", "")
                .trim();
        if (pattern.isEmpty() || "General".equalsIgnoreCase(pattern)) {
            return null;
        }
        if (pattern.contains("%")) {
            return "0%";
        }
        return pattern.matches(".*[0#].*") ? pattern : null;
    }

    private static List<String> orderedCategories(ChartSeries series, boolean reverse) {
        List<String> categories = new ArrayList<>(series.categories());
        if (reverse) {
            Collections.reverse(categories);
        }
        return categories;
    }

    private static List<Double> boxedValues(double[] values, boolean reverse, boolean zeroForNaN) {
        List<Double> boxed = new ArrayList<>(values.length);
        for (int i = 0; i < values.length; i++) {
            double value = values[reverse ? values.length - 1 - i : i];
            boxed.add(zeroForNaN && Double.isNaN(value) ? 0 : value);
        }
        return boxed;
    }

    private static double[] totals(List<ChartSeries> series) {
        int count = 0;
        for (ChartSeries item : series) {
            count = Math.max(count, item.values().length);
        }
        double[] totals = new double[count];
        for (ChartSeries item : series) {
            for (int i = 0; i < item.values().length; i++) {
                double value = item.values()[i];
                if (!Double.isNaN(value)) {
                    totals[i] += value;
                }
            }
        }
        for (int i = 0; i < totals.length; i++) {
            if (totals[i] == 0) {
                totals[i] = 1;
            }
        }
        return totals;
    }

    private static double[] percentValues(double[] values, double[] totals) {
        double[] result = new double[values.length];
        for (int i = 0; i < values.length; i++) {
            result[i] = Double.isNaN(values[i]) ? 0 : values[i] / totals[i];
        }
        return result;
    }

    private static Color sliceColor(DrawingTheme theme, int index) {
        Color color = theme.color("accent" + (index % 6 + 1));
        if (color == null) {
            color = new Color(0x4F, 0x81, 0xBD);
        }
        return color;
    }
}
