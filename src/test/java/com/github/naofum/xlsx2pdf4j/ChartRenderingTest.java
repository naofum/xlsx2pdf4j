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

package com.github.naofum.xlsx2pdf4j;

import com.github.naofum.xlsx2pdf4j.api.ConversionOptions;
import com.github.naofum.xlsx2pdf4j.api.ConversionResult;
import com.github.naofum.xlsx2pdf4j.api.Xlsx2Pdf;
import com.github.naofum.xlsx2pdf4j.api.TextLayerMode;
import com.github.naofum.xlsx2pdf4j.internal.chart.ChartKind;
import com.github.naofum.xlsx2pdf4j.internal.chart.ChartSpec;
import com.github.naofum.xlsx2pdf4j.internal.chart.XlsxChartReader;
import com.github.naofum.xlsx2pdf4j.internal.drawingml.DrawingTheme;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.BarDirection;
import org.apache.poi.xddf.usermodel.chart.BarGrouping;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.Grouping;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.awt.image.BufferedImage;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Exercises the XChart bridge with a workbook containing column, line, pie and
 * horizontal bar charts. The chart data is read from cached cell ranges, as in
 * real files.
 */
class ChartRenderingTest {

    @Test
    void rendersColumnLinePieAndHorizontalBarChartsWithoutWarnings() throws Exception {
        Path workbookPath = Path.of("target", "chart-rendering.xlsx");
        Path pdfPath = Path.of("target", "chart-rendering.pdf");
        Files.createDirectories(workbookPath.getParent());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet data = workbook.createSheet("Data");
            data.createRow(0).createCell(0).setCellValue("Quarter");
            data.getRow(0).createCell(1).setCellValue("2025");
            data.getRow(0).createCell(2).setCellValue("2026");
            String[] quarters = {"Q1", "Q2", "Q3", "Q4"};
            double[] values2025 = {120, 180, 150, 240};
            double[] values2026 = {90, 160, 210, 280};
            for (int i = 0; i < quarters.length; i++) {
                var row = data.createRow(i + 1);
                row.createCell(0).setCellValue(quarters[i]);
                row.createCell(1).setCellValue(values2025[i]);
                row.createCell(2).setCellValue(values2026[i]);
            }
            workbook.setSheetHidden(workbook.getSheetIndex(data), true);

            XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                    data, new CellRangeAddress(1, 4, 0, 0));
            XDDFNumericalDataSource<Double> series2025 = XDDFDataSourcesFactory.fromNumericCellRange(
                    data, new CellRangeAddress(1, 4, 1, 1));
            XDDFNumericalDataSource<Double> series2026 = XDDFDataSourcesFactory.fromNumericCellRange(
                    data, new CellRangeAddress(1, 4, 2, 2));

            XSSFSheet columnSheet = chartSheet(workbook, "Column");
            XSSFChart column = createChart(columnSheet);
            XDDFCategoryAxis categoryAxis = column.createCategoryAxis(AxisPosition.BOTTOM);
            XDDFValueAxis valueAxis = column.createValueAxis(AxisPosition.LEFT);
            XDDFBarChartData columnData = (XDDFBarChartData) column.createData(
                    ChartTypes.BAR, categoryAxis, valueAxis);
            columnData.setBarDirection(BarDirection.COL);
            columnData.setBarGrouping(BarGrouping.CLUSTERED);
            columnData.addSeries(categories, series2025).setTitle("2025", null);
            columnData.addSeries(categories, series2026).setTitle("2026", null);
            column.plot(columnData);
            column.setTitleText("Quarterly sales");
            column.getOrAddLegend().setPosition(LegendPosition.BOTTOM);

            XSSFSheet lineSheet = chartSheet(workbook, "Line");
            XSSFChart line = createChart(lineSheet);
            XDDFCategoryAxis lineCategoryAxis = line.createCategoryAxis(AxisPosition.BOTTOM);
            XDDFValueAxis lineValueAxis = line.createValueAxis(AxisPosition.LEFT);
            XDDFLineChartData lineData = (XDDFLineChartData) line.createData(
                    ChartTypes.LINE, lineCategoryAxis, lineValueAxis);
            lineData.setGrouping(Grouping.STANDARD);
            lineData.addSeries(categories, series2025).setTitle("2025", null);
            lineData.addSeries(categories, series2026).setTitle("2026", null);
            line.plot(lineData);
            line.setTitleText("Quarterly trend");

            XSSFSheet pieSheet = chartSheet(workbook, "Pie");
            XSSFChart pie = createChart(pieSheet);
            XDDFPieChartData pieData = (XDDFPieChartData) pie.createData(
                    ChartTypes.PIE, null, null);
            pieData.addSeries(categories, series2025);
            pie.plot(pieData);
            pie.setTitleText("Quarterly share");

            XSSFSheet barSheet = chartSheet(workbook, "Bar");
            XSSFChart bar = createChart(barSheet);
            XDDFCategoryAxis barCategoryAxis = bar.createCategoryAxis(AxisPosition.LEFT);
            XDDFValueAxis barValueAxis = bar.createValueAxis(AxisPosition.BOTTOM);
            XDDFBarChartData barData = (XDDFBarChartData) bar.createData(
                    ChartTypes.BAR, barCategoryAxis, barValueAxis);
            barData.setBarDirection(BarDirection.BAR);
            barData.setBarGrouping(BarGrouping.CLUSTERED);
            barData.addSeries(categories, series2025).setTitle("2025", null);
            barData.addSeries(categories, series2026).setTitle("2026", null);
            bar.plot(barData);
            bar.setTitleText("Quarterly sales by quarter");
            bar.getOrAddLegend().setPosition(LegendPosition.BOTTOM);

            ChartSpec columnSpec = XlsxChartReader.read(column, DrawingTheme.defaults());
            assertThat(columnSpec).isNotNull();
            assertThat(columnSpec.kind()).isEqualTo(ChartKind.COLUMN);
            assertThat(columnSpec.series()).hasSize(2);
            assertThat(columnSpec.series().get(0).name()).isEqualTo("2025");
            assertThat(columnSpec.series().get(0).values()).containsExactly(120, 180, 150, 240);
            assertThat(columnSpec.legendVisible()).isTrue();

            ChartSpec lineSpec = XlsxChartReader.read(line, DrawingTheme.defaults());
            assertThat(lineSpec).isNotNull();
            assertThat(lineSpec.kind()).isEqualTo(ChartKind.LINE);
            assertThat(lineSpec.series()).hasSize(2);

            ChartSpec pieSpec = XlsxChartReader.read(pie, DrawingTheme.defaults());
            assertThat(pieSpec).isNotNull();
            assertThat(pieSpec.kind()).isEqualTo(ChartKind.PIE);
            assertThat(pieSpec.series()).hasSize(1);

            ChartSpec barSpec = XlsxChartReader.read(bar, DrawingTheme.defaults());
            assertThat(barSpec).isNotNull();
            assertThat(barSpec.kind()).isEqualTo(ChartKind.BAR);
            assertThat(barSpec.horizontal()).isTrue();
            assertThat(barSpec.categoryReversed()).isFalse();
            assertThat(barSpec.series()).hasSize(2);
            assertThat(barSpec.series().get(0).values()).containsExactly(120, 180, 150, 240);
            assertThat(barSpec.legendVisible()).isTrue();

            try (var out = Files.newOutputStream(workbookPath)) {
                workbook.write(out);
            }
        }

        ConversionResult result = Xlsx2Pdf.convert(workbookPath, pdfPath,
                ConversionOptions.builder()
                        .dpi(150)
                        .textLayerMode(TextLayerMode.NONE)
                        .build());
        assertThat(result.pageCount()).isEqualTo(4);
        assertThat(result.warnings()).isEmpty();

        try (PDDocument document = Loader.loadPDF(pdfPath.toFile())) {
            PDFRenderer renderer = new PDFRenderer(document);
            for (int i = 0; i < 4; i++) {
                BufferedImage image = renderer.renderImageWithDPI(i, 150, ImageType.RGB);
                assertThat(nonWhitePixels(image))
                        .as("chart page %s", i + 1)
                        .isGreaterThan(500);
            }
        }
    }

    private static XSSFSheet chartSheet(XSSFWorkbook workbook, String name) {
        XSSFSheet sheet = workbook.createSheet(name);
        for (int c = 0; c < 8; c++) {
            sheet.setColumnWidth(c, 12 * 256);
        }
        for (int r = 0; r < 20; r++) {
            if (sheet.getRow(r) == null) {
                sheet.createRow(r);
            }
            sheet.getRow(r).setHeightInPoints(20);
        }
        workbook.setPrintArea(workbook.getSheetIndex(sheet), "A1:H20");
        return sheet;
    }

    private static XSSFChart createChart(XSSFSheet sheet) {
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
        return drawing.createChart(new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 7, 18));
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
