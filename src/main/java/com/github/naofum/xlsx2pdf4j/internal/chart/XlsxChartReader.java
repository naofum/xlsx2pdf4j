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

import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingStyle;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingTheme;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.chart.AxisOrientation;
import org.apache.poi.xddf.usermodel.chart.BarDirection;
import org.apache.poi.xddf.usermodel.chart.BarGrouping;
import org.apache.poi.xddf.usermodel.chart.Grouping;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFAreaChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChart;
import org.apache.poi.xddf.usermodel.chart.XDDFTitle;
import org.apache.poi.xddf.usermodel.text.XDDFTextParagraph;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;

import java.awt.Color;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.IdentityHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

/**
 * Reads an {@link XDDFChart} (spreadsheet or presentation) into the
 * format-neutral chart model.
 */
public final class XlsxChartReader {

    private XlsxChartReader() {
    }

    public static ChartSpec read(XDDFChart chart, DrawingTheme theme) {
        if (chart == null) {
            return null;
        }
        List<XDDFChartData> groups = chart.getChartSeries();
        if (groups == null || groups.isEmpty()) {
            return null;
        }
        DrawingTheme colors = theme == null ? DrawingTheme.defaults() : theme;
        List<XDDFValueAxis> valueAxes = new ArrayList<>();
        Map<XDDFValueAxis, Integer> axisIndexes = new IdentityHashMap<>();
        List<ChartSeries> series = new ArrayList<>();
        int seriesIndex = 0;

        ChartKind chartKind = null;
        boolean stacked = false;
        boolean percentStacked = false;
        boolean horizontal = false;
        for (XDDFChartData group : groups) {
            ChartKind groupKind = kind(group);
            if (groupKind == null) {
                continue;
            }
            if (chartKind == null) {
                chartKind = groupKind;
                stacked = stacked(group);
                percentStacked = percentStacked(group);
                horizontal = groupKind == ChartKind.BAR;
            }
            int axisGroup = 0;
            List<XDDFValueAxis> groupAxes = group.getValueAxes();
            if (groupAxes != null && !groupAxes.isEmpty()) {
                XDDFValueAxis axis = groupAxes.get(0);
                Integer existing = axisIndexes.get(axis);
                if (existing == null) {
                    existing = valueAxes.size();
                    valueAxes.add(axis);
                    axisIndexes.put(axis, existing);
                }
                axisGroup = existing;
            }
            for (int i = 0; i < group.getSeriesCount(); i++) {
                ChartSeries converted = convertSeries(group.getSeries(i), groupKind,
                        seriesIndex, axisGroup, colors);
                if (converted != null) {
                    series.add(converted);
                    seriesIndex++;
                }
            }
        }
        if (chartKind == null || series.isEmpty()) {
            return null;
        }

        ChartAxis categoryAxis = categoryAxis(groups.get(0));
        boolean categoryReversed = false;
        XDDFCategoryAxis categoryAxisObject = groups.get(0).getCategoryAxis();
        if (categoryAxisObject != null
                && categoryAxisObject.getOrientation() == AxisOrientation.MAX_MIN) {
            categoryReversed = true;
        }
        List<ChartAxis> axisModels = new ArrayList<>();
        for (XDDFValueAxis axis : valueAxes) {
            axisModels.add(valueAxis(axis));
        }
        if (axisModels.isEmpty()) {
            axisModels.add(new ChartAxis(null, null, null, null, null, false, true));
        }

        boolean legendVisible = chart.getCTChart() != null
                && chart.getCTChart().isSetLegend();
        String legendPosition = "right";
        if (legendVisible) {
            XDDFChartLegend legend = new XDDFChartLegend(chart.getCTChart());
            LegendPosition position = legend.getPosition();
            if (position != null) {
                legendPosition = position.name().toLowerCase(Locale.ROOT);
            }
        }
        String title = titleText(chart);
        double firstSliceAngle = 0;
        double holeSize = 0;
        for (XDDFChartData group : groups) {
            if (group instanceof XDDFPieChartData pie && pie.getFirstSliceAngle() != null) {
                firstSliceAngle = pie.getFirstSliceAngle();
            }
            if (group instanceof XDDFDoughnutChartData doughnut
                    && doughnut.getFirstSliceAngle() != null) {
                firstSliceAngle = doughnut.getFirstSliceAngle();
            }
            if (group instanceof XDDFDoughnutChartData doughnut
                    && doughnut.getHoleSize() != null) {
                holeSize = doughnut.getHoleSize();
            }
        }
        return new ChartSpec(chartKind, title, legendVisible, legendPosition, series,
                categoryAxis, categoryReversed, axisModels, horizontal, stacked, percentStacked,
                firstSliceAngle, holeSize);
    }

    private static String titleText(XDDFChart chart) {
        try {
            XDDFTitle title = chart.getTitle();
            if (title == null || title.getBody() == null) {
                return null;
            }
            StringBuilder text = new StringBuilder();
            for (XDDFTextParagraph paragraph : title.getBody().getParagraphs()) {
                text.append(paragraph.getText());
            }
            return text.length() == 0 ? null : text.toString();
        } catch (RuntimeException ignored) {
            return null;
        }
    }

    private static ChartSeries convertSeries(XDDFChartData.Series series, ChartKind kind,
                                              int index, int axisGroup,
                                              DrawingTheme theme) {
        XDDFNumericalDataSource<? extends Number> valuesData = series.getValuesData();
        int count = valuesData == null ? 0 : valuesData.getPointCount();
        if (count == 0) {
            return null;
        }
        double[] values = new double[count];
        for (int i = 0; i < count; i++) {
            Number value = valuesData.getPointAt(i);
            values[i] = value == null ? Double.NaN : value.doubleValue();
        }
        XDDFDataSource<?> categoryData = series.getCategoryData();
        List<String> categories = new ArrayList<>(count);
        for (int i = 0; i < count; i++) {
            Object category = categoryData != null && i < categoryData.getPointCount()
                    ? categoryData.getPointAt(i) : null;
            categories.add(category == null ? String.valueOf(i + 1) : String.valueOf(category));
        }

        Color fill = theme.color("accent" + (index % 6 + 1));
        if (fill == null) {
            fill = Color.GRAY;
        }
        Color line = fill.darker();
        double lineWidth = 1.0;
        XDDFShapeProperties properties = series.getShapeProperties();
        if (properties != null && properties.getXmlObject() != null) {
            DrawingStyle style = DrawingStyle.parse(
                    properties.getXmlObject().xmlText(), theme.colors());
            if (style.fill != null) {
                fill = style.fill;
            }
            if (style.line != null) {
                line = style.line;
            }
            if (style.lineWidth > 0) {
                lineWidth = style.lineWidth;
            }
        }
        return new ChartSeries(seriesName(series, index), categories, values,
                axisGroup, kind, fill, line, lineWidth, null, false);
    }

    private static String seriesName(XDDFChartData.Series series, int index) {
        try {
            Method method = XDDFChartData.Series.class.getDeclaredMethod("getSeriesText");
            method.setAccessible(true);
            CTSerTx text = (CTSerTx) method.invoke(series);
            if (text != null) {
                if (text.isSetV() && !text.getV().isBlank()) {
                    return text.getV();
                }
                if (text.isSetStrRef() && text.getStrRef().isSetStrCache()
                        && !text.getStrRef().getStrCache().getPtList().isEmpty()) {
                    return text.getStrRef().getStrCache().getPtList().get(0).getV();
                }
            }
        } catch (ReflectiveOperationException ignored) {
            // Fall back to a stable generated name.
        }
        return "Series " + (index + 1);
    }

    private static ChartKind kind(XDDFChartData data) {
        if (data instanceof XDDFBarChartData bar) {
            return bar.getBarDirection() == BarDirection.BAR
                    ? ChartKind.BAR : ChartKind.COLUMN;
        }
        if (data instanceof XDDFLineChartData) {
            return ChartKind.LINE;
        }
        if (data instanceof XDDFDoughnutChartData) {
            return ChartKind.DOUGHNUT;
        }
        if (data instanceof XDDFPieChartData) {
            return ChartKind.PIE;
        }
        if (data instanceof XDDFAreaChartData) {
            return ChartKind.AREA;
        }
        if (data instanceof XDDFScatterChartData) {
            return ChartKind.SCATTER;
        }
        return null;
    }

    private static boolean stacked(XDDFChartData data) {
        if (data instanceof XDDFBarChartData bar) {
            BarGrouping grouping = bar.getBarGrouping();
            return grouping == BarGrouping.STACKED || grouping == BarGrouping.PERCENT_STACKED;
        }
        if (data instanceof XDDFLineChartData line) {
            Grouping grouping = line.getGrouping();
            return grouping == Grouping.STACKED || grouping == Grouping.PERCENT_STACKED;
        }
        if (data instanceof XDDFAreaChartData area) {
            Grouping grouping = area.getGrouping();
            return grouping == Grouping.STACKED || grouping == Grouping.PERCENT_STACKED;
        }
        return false;
    }

    private static boolean percentStacked(XDDFChartData data) {
        if (data instanceof XDDFBarChartData bar) {
            return bar.getBarGrouping() == BarGrouping.PERCENT_STACKED;
        }
        if (data instanceof XDDFLineChartData line) {
            return line.getGrouping() == Grouping.PERCENT_STACKED;
        }
        if (data instanceof XDDFAreaChartData area) {
            return area.getGrouping() == Grouping.PERCENT_STACKED;
        }
        return false;
    }

    private static ChartAxis categoryAxis(XDDFChartData data) {
        XDDFCategoryAxis axis = data.getCategoryAxis();
        if (axis == null) {
            return new ChartAxis(null, null, null, null, null, false, true);
        }
        return new ChartAxis(null, null, null, null, numberFormat(axis),
                false, axis.isVisible());
    }

    private static ChartAxis valueAxis(XDDFValueAxis axis) {
        if (axis == null) {
            return new ChartAxis(null, null, null, null, null, false, true);
        }
        return new ChartAxis(null,
                axis.isSetMinimum() ? axis.getMinimum() : null,
                axis.isSetMaximum() ? axis.getMaximum() : null,
                axis.isSetMajorUnit() ? axis.getMajorUnit() : null,
                numberFormat(axis),
                axis.isSetLogBase(),
                axis.isVisible());
    }

    private static String numberFormat(XDDFChartAxis axis) {
        return axis != null && axis.hasNumberFormat() ? axis.getNumberFormat() : null;
    }
}
