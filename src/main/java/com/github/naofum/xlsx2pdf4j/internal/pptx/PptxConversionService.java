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

import com.github.naofum.xlsx2pdf4j.api.ConversionOptions;
import com.github.naofum.xlsx2pdf4j.api.ConversionResult;
import com.github.naofum.xlsx2pdf4j.api.ConversionWarning;
import com.github.naofum.xlsx2pdf4j.api.ErrorCode;
import com.github.naofum.xlsx2pdf4j.api.ImageCompression;
import com.github.naofum.xlsx2pdf4j.api.TextLayerMode;
import com.github.naofum.xlsx2pdf4j.api.WarningCode;
import com.github.naofum.xlsx2pdf4j.api.Xlsx2PdfException;
import com.github.naofum.xlsx2pdf4j.internal.FontResolver;
import com.github.naofum.xlsx2pdf4j.internal.PdfWriter;
import com.github.naofum.xlsx2pdf4j.internal.PlacedText;
import com.github.naofum.xlsx2pdf4j.internal.chart.ChartRenderer;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Dimension;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

/**
 * Orchestrates presentation reading, slide rendering and PDF writing.
 */
public final class PptxConversionService {

    private static final Logger LOG = LoggerFactory.getLogger(PptxConversionService.class);
    private static final long MAX_INPUT_BYTES = 100L * 1024 * 1024;

    public ConversionResult convert(Path input, Path output, ConversionOptions options)
            throws Xlsx2PdfException {
        Instant started = Instant.now();
        validateInput(input);
        if (output == null) {
            throw new Xlsx2PdfException(ErrorCode.OUTPUT_IO, "output path is null");
        }
        preparePoiSecurity();

        List<ConversionWarning> warnings = new ArrayList<>();
        Path absoluteOutput = output.toAbsolutePath();
        try (OPCPackage pkg = OPCPackage.open(input.toFile(), PackageAccess.READ);
             XMLSlideShow presentation = new XMLSlideShow(pkg)) {
            List<XSLFSlide> slides = selectSlides(presentation, options);
            if (slides.isEmpty()) {
                throw new Xlsx2PdfException(ErrorCode.INPUT_FORMAT,
                        "no printable slides were selected");
            }
            if (slides.size() > options.maxPages()) {
                throw new Xlsx2PdfException(ErrorCode.LIMIT_EXCEEDED,
                        "slide count " + slides.size() + " exceeds limit " + options.maxPages());
            }

            PptxFontNormalizer.normalize(presentation);
            FontResolver fonts = new FontResolver();
            ChartRenderer charts = new ChartRenderer(fonts);
            ImageCompression compression = effectiveImageCompression(options);
            Dimension pageSize = presentation.getPageSize();
            String fileName = input.getFileName() == null ? "" : input.getFileName().toString();

            try (PdfWriter pdf = new PdfWriter(absoluteOutput, options.maxOutputBytes(),
                    compression, options.jpegQuality())) {
                Map<String, Integer> pagesBySheet = new LinkedHashMap<>();
                int pageNumber = 1;
                for (XSLFSlide slide : slides) {
                    List<PlacedText> textLayer = options.textLayerMode() == TextLayerMode.INVISIBLE
                            ? new ArrayList<>() : null;
                    BufferedImage image = PptxSlideRenderer.render(presentation, slide,
                            options.dpi(), charts, textLayer, warnings);
                    checkStrict(options, warnings);
                    pdf.addPage(image, pageSize.getWidth(), pageSize.getHeight(), textLayer);
                    image.flush();
                    pagesBySheet.put("Slide " + pageNumber, 1);
                    pageNumber++;
                }
                pdf.finish(fileName);
                long size = Files.size(absoluteOutput);
                LOG.info("converted {} slides, {} bytes", slides.size(), size);
                return new ConversionResult(absoluteOutput, slides.size(), pagesBySheet, warnings,
                        Duration.between(started, Instant.now()), size);
            }
        } catch (Xlsx2PdfException e) {
            throw e;
        } catch (EncryptedDocumentException e) {
            throw new Xlsx2PdfException(ErrorCode.INPUT_ENCRYPTED,
                    "the presentation is encrypted", e);
        } catch (InvalidFormatException e) {
            throw new Xlsx2PdfException(ErrorCode.INPUT_FORMAT,
                    "invalid .pptx package", e);
        } catch (IOException e) {
            throw new Xlsx2PdfException(ErrorCode.INPUT_CORRUPT,
                    "cannot read presentation", e);
        } catch (RuntimeException e) {
            throw new Xlsx2PdfException(ErrorCode.INTERNAL,
                    "unexpected presentation conversion error", e);
        }
    }

    private List<XSLFSlide> selectSlides(XMLSlideShow presentation, ConversionOptions options) {
        Set<String> requested = options.sheets();
        List<XSLFSlide> slides = new ArrayList<>();
        int index = 1;
        for (XSLFSlide slide : presentation.getSlides()) {
            boolean selected = requested.isEmpty()
                    || requested.contains(Integer.toString(index))
                    || requested.contains("Slide " + index)
                    || (slide.getTitle() != null && requested.contains(slide.getTitle()));
            if (selected && (!slide.isHidden() || !requested.isEmpty())) {
                slides.add(slide);
            }
            index++;
        }
        return slides;
    }

    private void checkStrict(ConversionOptions options, List<ConversionWarning> warnings)
            throws Xlsx2PdfException {
        if (!options.strict()) {
            return;
        }
        for (ConversionWarning warning : warnings) {
            if (warning.code() == WarningCode.SHAPE_UNSUPPORTED
                    || warning.code() == WarningCode.IMAGE_UNSUPPORTED
                    || warning.code() == WarningCode.CHART_SKIPPED
                    || warning.code() == WarningCode.SMART_ART_SKIPPED) {
                throw new Xlsx2PdfException(ErrorCode.UNSUPPORTED,
                        "strict mode: " + warning.message(),
                        warning.sheetName(), warning.location());
            }
        }
    }

    private ImageCompression effectiveImageCompression(ConversionOptions options) {
        if (options.imageCompression() == ImageCompression.AUTO && options.dpi() >= 300) {
            return ImageCompression.JPEG;
        }
        return options.imageCompression();
    }

    private void validateInput(Path input) throws Xlsx2PdfException {
        if (input == null) {
            throw new Xlsx2PdfException(ErrorCode.INPUT_NOT_FOUND, "input path is null");
        }
        if (!Files.isRegularFile(input)) {
            throw new Xlsx2PdfException(ErrorCode.INPUT_NOT_FOUND,
                    "input file not found: " + input);
        }
        String name = input.getFileName() == null
                ? "" : input.getFileName().toString().toLowerCase(Locale.ROOT);
        if (!name.endsWith(".pptx") && !name.endsWith(".pptm")) {
            throw new Xlsx2PdfException(ErrorCode.INPUT_FORMAT,
                    "only .pptx/.pptm is supported: " + input);
        }
        try {
            if (Files.size(input) > MAX_INPUT_BYTES) {
                throw new Xlsx2PdfException(ErrorCode.LIMIT_EXCEEDED,
                        "input file is larger than 100MB: " + input);
            }
        } catch (IOException e) {
            throw new Xlsx2PdfException(ErrorCode.INPUT_NOT_FOUND,
                    "cannot stat input file", e);
        }
    }

    private void preparePoiSecurity() {
        ZipSecureFile.setMinInflateRatio(0.001d);
        ZipSecureFile.setMaxEntrySize(MAX_INPUT_BYTES);
        ZipSecureFile.setMaxTextSize(MAX_INPUT_BYTES);
    }
}
