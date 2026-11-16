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

import com.github.naofum.xlsx2pdf4j.api.Pptx2Pdf;
import com.github.naofum.xlsx2pdf4j.api.PptxConversionOptions;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.junit.jupiter.api.Test;

import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Coarse regression guard against PowerPoint-generated reference PDFs.
 */
class PptxGoldenPdfComparisonTest {

    private static final double MAX_MEAN_DIFFERENCE = 0.05;
    private static final boolean WRITE_DIFF_IMAGES =
            !"false".equalsIgnoreCase(System.getProperty("xlsx2pdf4j.diffImages", "true"));

    @Test
    void sample1IsCloseToGoldenPdf() throws Exception {
        compare("sample1", 1);
    }

    @Test
    void sample2IsCloseToGoldenPdf() throws Exception {
        compare("sample2", 3);
    }

    @Test
    void sample3IsCloseToGoldenPdf() throws Exception {
        compare("sample3", 1);
    }

    private void compare(String name, int expectedPages) throws Exception {
        Path workbook = Path.of("testdata", "validation", name + ".pptx");
        Path candidatePath = Path.of("target", "golden-candidate", "pptx-" + name + ".pdf");
        Path goldenPath = Path.of("testdata", "golden", name + ".pdf");
        Files.createDirectories(candidatePath.getParent());

        Pptx2Pdf.convert(workbook, candidatePath,
                PptxConversionOptions.builder().dpi(150).build());

        assertThat(candidatePath).exists();
        assertThat(goldenPath).exists();
        Path diffDirectory = Path.of("target", "golden-diff", "pptx-" + name);
        if (WRITE_DIFF_IMAGES) {
            Files.createDirectories(diffDirectory);
        }
        try (PDDocument candidate = Loader.loadPDF(candidatePath.toFile());
             PDDocument golden = Loader.loadPDF(goldenPath.toFile())) {
            assertThat(candidate.getNumberOfPages()).isEqualTo(expectedPages);
            assertThat(golden.getNumberOfPages()).isEqualTo(expectedPages);

            PDFRenderer candidateRenderer = new PDFRenderer(candidate);
            PDFRenderer goldenRenderer = new PDFRenderer(golden);
            double[] differences = new double[expectedPages];
            double total = 0;
            for (int i = 0; i < expectedPages; i++) {
                BufferedImage a = candidateRenderer.renderImageWithDPI(i, 72, ImageType.RGB);
                BufferedImage b = goldenRenderer.renderImageWithDPI(i, 72, ImageType.RGB);
                differences[i] = meanDifference(a, b);
                total += differences[i];
                if (WRITE_DIFF_IMAGES) {
                    PageDiffWriter.write(diffDirectory, i + 1, a, b);
                }
            }
            double mean = total / expectedPages;
            System.out.printf("%s mean image difference = %.4f%n", name, mean);
            if (WRITE_DIFF_IMAGES) {
                PageDiffWriter.writeSummary(diffDirectory, differences);
                assertThat(diffDirectory.resolve("page-01-diff.png")).exists();
                assertThat(diffDirectory.resolve(
                        String.format(java.util.Locale.ROOT, "page-%02d-diff.png", expectedPages)))
                        .exists();
                assertThat(diffDirectory.resolve("summary.txt")).exists();
            }
            assertThat(mean).isLessThan(MAX_MEAN_DIFFERENCE);
        }
    }

    private double meanDifference(BufferedImage a, BufferedImage b) {
        if (a.getWidth() != b.getWidth() || a.getHeight() != b.getHeight()) {
            b = resize(b, a.getWidth(), a.getHeight());
        }
        long sum = 0;
        long count = 0;
        for (int y = 0; y < a.getHeight(); y++) {
            for (int x = 0; x < a.getWidth(); x++) {
                int ca = a.getRGB(x, y);
                int cb = b.getRGB(x, y);
                sum += Math.abs(((ca >> 16) & 0xFF) - ((cb >> 16) & 0xFF));
                sum += Math.abs(((ca >> 8) & 0xFF) - ((cb >> 8) & 0xFF));
                sum += Math.abs((ca & 0xFF) - (cb & 0xFF));
                count += 3;
            }
        }
        return count == 0 ? 1 : (sum / (double) count) / 255d;
    }

    private BufferedImage resize(BufferedImage source, int width, int height) {
        BufferedImage result = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g = result.createGraphics();
        try {
            g.setRenderingHint(RenderingHints.KEY_INTERPOLATION,
                    RenderingHints.VALUE_INTERPOLATION_BILINEAR);
            g.drawImage(source, 0, 0, width, height, null);
        } finally {
            g.dispose();
        }
        return result;
    }
}
