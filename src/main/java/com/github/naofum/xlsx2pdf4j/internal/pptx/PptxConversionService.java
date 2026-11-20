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

import com.github.naofum.xlsx2pdf4j.api.ConversionWarning;
import com.github.naofum.xlsx2pdf4j.api.ErrorCode;
import com.github.naofum.xlsx2pdf4j.api.ImageCompression;
import com.github.naofum.xlsx2pdf4j.api.Pptx2PdfException;
import com.github.naofum.xlsx2pdf4j.api.PptxConversionOptions;
import com.github.naofum.xlsx2pdf4j.api.PptxConversionResult;
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

    public PptxConversionResult convert(Path input, Path output, PptxConversionOptions options)
            throws Pptx2PdfException {
        Instant started = Instant.now();
        validateInput(input);
        if (output == null) {
            throw new Pptx2PdfException(ErrorCode.OUTPUT_IO, "output path is null");
        }
        preparePoiSecurity();

        List<ConversionWarning> warnings = new ArrayList<>();
        Path absoluteOutput = output.toAbsolutePath();
        try (OPCPackage pkg = OPCPackage.open(input.toFile(), PackageAccess.READ);
             XMLSlideShow presentation = new XMLSlideShow(pkg)) {
            Map<Integer, XSLFSlide> selected = selectSlides(presentation, options);
            if (selected.isEmpty()) {
                throw new Pptx2PdfException(ErrorCode.INPUT_FORMAT,
                        "no printable slides were selected");
            }
            if (selected.size() > options.maxSlides()) {
                throw new Pptx2PdfException(ErrorCode.LIMIT_EXCEEDED,
                        "slide count " + selected.size() + " exceeds limit " + options.maxSlides());
            }

            PptxFontNormalizer.normalize(presentation);
            FontResolver fonts = new FontResolver();
            ChartRenderer charts = new ChartRenderer(fonts);
            ImageCompression compression = effectiveImageCompression(options);
            Dimension pageSize = presentation.getPageSize();
            String fileName = input.getFileName() == null ? "" : input.getFileName().toString();

            try (PdfWriter pdf = new PdfWriter(absoluteOutput, options.maxOutputBytes(),
                    compression, options.jpegQuality(), options.pdfMetadata())) {
                int physicalPage = 1;
                for (XSLFSlide slide : selected.values()) {
                    List<PlacedText> textLayer = options.textLayerMode() == TextLayerMode.INVISIBLE
                            ? new ArrayList<>() : null;
                    BufferedImage image = PptxSlideRenderer.render(presentation, slide,
                            options.dpi(), charts, textLayer, warnings);
                    checkStrict(options, warnings);
                    pdf.addPage(image, pageSize.getWidth(), pageSize.getHeight(), textLayer);
                    if (options.slideBookmarks()) {
                        String title = slide.getTitle();
                        if (title != null && !title.isBlank()) {
                            pdf.addBookmark(title.strip(), physicalPage);
                        }
                    }
                    physicalPage++;
                    image.flush();
                }
                pdf.finish(fileName);
                long size = Files.size(absoluteOutput);
                LOG.info("converted {} slides, {} bytes", selected.size(), size);
                return new PptxConversionResult(absoluteOutput,
                        new ArrayList<>(selected.keySet()), warnings,
                        Duration.between(started, Instant.now()), size);
            }
        } catch (Pptx2PdfException e) {
            throw e;
        } catch (Xlsx2PdfException e) {
            throw new Pptx2PdfException(e.code(), e.getMessage(), e);
        } catch (EncryptedDocumentException e) {
            throw new Pptx2PdfException(ErrorCode.INPUT_ENCRYPTED,
                    "the presentation is encrypted", e);
        } catch (InvalidFormatException e) {
            throw new Pptx2PdfException(ErrorCode.INPUT_FORMAT,
                    "invalid .pptx package", e);
        } catch (IOException e) {
            throw new Pptx2PdfException(ErrorCode.INPUT_CORRUPT,
                    "cannot read presentation", e);
        } catch (RuntimeException e) {
            throw new Pptx2PdfException(ErrorCode.INTERNAL,
                    "unexpected presentation conversion error", e);
        }
    }

    private Map<Integer, XSLFSlide> selectSlides(XMLSlideShow presentation,
                                                 PptxConversionOptions options) {
        Set<Integer> requested = options.slides();
        Map<Integer, XSLFSlide> slides = new LinkedHashMap<>();
        int index = 1;
        for (XSLFSlide slide : presentation.getSlides()) {
            boolean selected = requested.isEmpty() || requested.contains(index);
            if (selected && (!slide.isHidden() || !requested.isEmpty())) {
                slides.put(index, slide);
            }
            index++;
        }
        return slides;
    }

    private void checkStrict(PptxConversionOptions options, List<ConversionWarning> warnings)
            throws Pptx2PdfException {
        if (!options.strict()) {
            return;
        }
        for (ConversionWarning warning : warnings) {
            if (warning.code() == WarningCode.SHAPE_UNSUPPORTED
                    || warning.code() == WarningCode.IMAGE_UNSUPPORTED
                    || warning.code() == WarningCode.CHART_SKIPPED
                    || warning.code() == WarningCode.SMART_ART_SKIPPED) {
                throw new Pptx2PdfException(ErrorCode.UNSUPPORTED,
                        "strict mode: " + warning.message(),
                        warning.sheetName(), warning.location());
            }
        }
    }

    private ImageCompression effectiveImageCompression(PptxConversionOptions options) {
        if (options.imageCompression() == ImageCompression.AUTO && options.dpi() >= 300) {
            return ImageCompression.JPEG;
        }
        return options.imageCompression();
    }

    private void validateInput(Path input) throws Pptx2PdfException {
        if (input == null) {
            throw new Pptx2PdfException(ErrorCode.INPUT_NOT_FOUND, "input path is null");
        }
        if (!Files.isRegularFile(input)) {
            throw new Pptx2PdfException(ErrorCode.INPUT_NOT_FOUND,
                    "input file not found: " + input);
        }
        String name = input.getFileName() == null
                ? "" : input.getFileName().toString().toLowerCase(Locale.ROOT);
        if (!name.endsWith(".pptx") && !name.endsWith(".pptm")) {
            throw new Pptx2PdfException(ErrorCode.INPUT_FORMAT,
                    "only .pptx/.pptm is supported: " + input);
        }
        try {
            if (Files.size(input) > MAX_INPUT_BYTES) {
                throw new Pptx2PdfException(ErrorCode.LIMIT_EXCEEDED,
                        "input file is larger than 100MB: " + input);
            }
        } catch (IOException e) {
            throw new Pptx2PdfException(ErrorCode.INPUT_NOT_FOUND,
                    "cannot stat input file", e);
        }
    }

    private void preparePoiSecurity() {
        ZipSecureFile.setMinInflateRatio(0.001d);
        ZipSecureFile.setMaxEntrySize(MAX_INPUT_BYTES);
        ZipSecureFile.setMaxTextSize(MAX_INPUT_BYTES);
    }
}
