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
import com.github.naofum.xlsx2pdf4j.api.Pptx2Pdf;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.junit.jupiter.api.Test;

import java.awt.geom.Rectangle2D;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;

class PptxConversionTest {

    @Test
    void convertsSamplePresentations() throws Exception {
        Path outputDir = Path.of("target", "pptx");
        Files.createDirectories(outputDir);
        for (String name : new String[]{"sample1", "sample2", "sample3"}) {
            ConversionResult result = Pptx2Pdf.convert(
                    Path.of("testdata", "validation", name + ".pptx"),
                    outputDir.resolve(name + ".pdf"),
                    ConversionOptions.builder().dpi(150).build());

            assertThat(result.pageCount()).isEqualTo(1);
            assertThat(result.pagesBySheet()).containsEntry("Slide 1", 1);
            assertThat(result.warnings()).isEmpty();
            assertThat(result.output()).exists();
            assertThat(result.outputSize()).isLessThan(50L * 1024 * 1024);
        }
    }

    @Test
    void writesSearchableTextLayer() throws Exception {
        Path outputDir = Path.of("target", "pptx");
        Files.createDirectories(outputDir);
        String[] expected = {"タイトル", "デモ画面", "Amazon Connect", "onConnecting"};
        String[] names = {"sample1", "sample1", "sample2", "sample3"};
        for (int i = 0; i < names.length; i++) {
            Path pdf = outputDir.resolve(names[i] + "-text.pdf");
            Pptx2Pdf.convert(Path.of("testdata", "validation", names[i] + ".pptx"), pdf,
                    ConversionOptions.builder().dpi(100).build());
            try (PDDocument document = Loader.loadPDF(pdf.toFile())) {
                String text = new PDFTextStripper().getText(document);
                assertThat(text).contains(expected[i]);
            }
        }
    }

    @Test
    void skipsHiddenSlidesUnlessSelected() throws Exception {
        Path workbook = Path.of("target", "pptx", "two-slides.pptx");
        Files.createDirectories(workbook.getParent());
        try (XMLSlideShow presentation = new XMLSlideShow()) {
            XSLFSlide first = presentation.createSlide();
            XSLFTextBox firstBox = first.createTextBox();
            firstBox.setAnchor(new Rectangle2D.Double(50, 50, 300, 80));
            firstBox.setText("First slide");
            XSLFSlide second = presentation.createSlide();
            XSLFTextBox secondBox = second.createTextBox();
            secondBox.setAnchor(new Rectangle2D.Double(50, 50, 300, 80));
            secondBox.setText("Second slide");
            second.setHidden(true);
            try (var out = Files.newOutputStream(workbook)) {
                presentation.write(out);
            }
        }

        ConversionResult visible = Pptx2Pdf.convert(workbook,
                Path.of("target", "pptx", "two-slides-visible.pdf"),
                ConversionOptions.builder().dpi(100).build());
        assertThat(visible.pageCount()).isEqualTo(1);

        ConversionResult selected = Pptx2Pdf.convert(workbook,
                Path.of("target", "pptx", "two-slides-selected.pdf"),
                ConversionOptions.builder().dpi(100).sheet("2").build());
        assertThat(selected.pageCount()).isEqualTo(1);
    }
}
