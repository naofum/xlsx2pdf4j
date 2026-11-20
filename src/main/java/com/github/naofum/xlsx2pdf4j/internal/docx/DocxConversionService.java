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

package com.github.naofum.xlsx2pdf4j.internal.docx;

import com.github.naofum.xlsx2pdf4j.api.ConversionWarning;
import com.github.naofum.xlsx2pdf4j.api.Docx2PdfException;
import com.github.naofum.xlsx2pdf4j.api.DocxConversionOptions;
import com.github.naofum.xlsx2pdf4j.api.DocxConversionResult;
import com.github.naofum.xlsx2pdf4j.api.ErrorCode;
import com.github.naofum.xlsx2pdf4j.api.ImageCompression;
import com.github.naofum.xlsx2pdf4j.api.WarningCode;
import com.github.naofum.xlsx2pdf4j.api.Xlsx2PdfException;
import com.github.naofum.xlsx2pdf4j.internal.PdfWriter;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Orchestrates Word reading, flow layout and PDF writing.
 */
public final class DocxConversionService {

    private static final Logger LOG = LoggerFactory.getLogger(DocxConversionService.class);
    private static final long MAX_INPUT_BYTES = 100L * 1024 * 1024;

    public DocxConversionResult convert(Path input, Path output, DocxConversionOptions options)
            throws Docx2PdfException {
        Instant started = Instant.now();
        validateInput(input);
        if (output == null) {
            throw new Docx2PdfException(ErrorCode.OUTPUT_IO, "output path is null");
        }
        preparePoiSecurity();
        List<ConversionWarning> warnings = new ArrayList<>();
        Path absoluteOutput = output.toAbsolutePath();
        try (OPCPackage pkg = OPCPackage.open(input.toFile(), PackageAccess.READ);
             XWPFDocument document = new XWPFDocument(pkg)) {
            DocxDocumentModel model = new DocxStyleResolver(document).resolve();
            ImageCompression compression = effectiveImageCompression(options);
            String fileName = input.getFileName() == null ? "" : input.getFileName().toString();
            try (PdfWriter pdf = new PdfWriter(absoluteOutput, options.maxOutputBytes(),
                    compression, options.jpegQuality(), options.pdfMetadata())) {
                DocxFlowLayout layout = new DocxFlowLayout(model, options, pdf, warnings);
                int pages = layout.render();
                checkStrict(options, warnings);
                pdf.finish(fileName);
                long size = Files.size(absoluteOutput);
                LOG.info("converted {} pages, {} bytes", pages, size);
                return new DocxConversionResult(absoluteOutput, pages, warnings,
                        Duration.between(started, Instant.now()), size);
            } catch (Xlsx2PdfException e) {
                throw new Docx2PdfException(e.code(), e.getMessage(), e);
            }
        } catch (Docx2PdfException e) {
            throw e;
        } catch (EncryptedDocumentException e) {
            throw new Docx2PdfException(ErrorCode.INPUT_ENCRYPTED,
                    "the document is encrypted", e);
        } catch (InvalidFormatException e) {
            throw new Docx2PdfException(ErrorCode.INPUT_FORMAT,
                    "invalid .docx package", e);
        } catch (IOException e) {
            throw new Docx2PdfException(ErrorCode.INPUT_CORRUPT,
                    "cannot read document", e);
        } catch (RuntimeException e) {
            throw new Docx2PdfException(ErrorCode.INTERNAL,
                    "unexpected document conversion error", e);
        }
    }

    private void checkStrict(DocxConversionOptions options, List<ConversionWarning> warnings)
            throws Docx2PdfException {
        if (!options.strict()) {
            return;
        }
        for (ConversionWarning warning : warnings) {
            if (warning.code() == WarningCode.SHAPE_UNSUPPORTED
                    || warning.code() == WarningCode.IMAGE_UNSUPPORTED) {
                throw new Docx2PdfException(ErrorCode.UNSUPPORTED,
                        "strict mode: " + warning.message(),
                        warning.sheetName(), warning.location());
            }
        }
    }

    private ImageCompression effectiveImageCompression(DocxConversionOptions options) {
        if (options.imageCompression() == ImageCompression.AUTO && options.dpi() >= 300) {
            return ImageCompression.JPEG;
        }
        return options.imageCompression();
    }

    private void validateInput(Path input) throws Docx2PdfException {
        if (input == null) {
            throw new Docx2PdfException(ErrorCode.INPUT_NOT_FOUND, "input path is null");
        }
        if (!Files.isRegularFile(input)) {
            throw new Docx2PdfException(ErrorCode.INPUT_NOT_FOUND,
                    "input file not found: " + input);
        }
        String name = input.getFileName() == null
                ? "" : input.getFileName().toString().toLowerCase(Locale.ROOT);
        if (!name.endsWith(".docx") && !name.endsWith(".docm")) {
            throw new Docx2PdfException(ErrorCode.INPUT_FORMAT,
                    "only .docx/.docm is supported: " + input);
        }
        try {
            if (Files.size(input) > MAX_INPUT_BYTES) {
                throw new Docx2PdfException(ErrorCode.LIMIT_EXCEEDED,
                        "input file is larger than 100MB: " + input);
            }
        } catch (IOException e) {
            throw new Docx2PdfException(ErrorCode.INPUT_NOT_FOUND,
                    "cannot stat input file", e);
        }
    }

    private void preparePoiSecurity() {
        ZipSecureFile.setMinInflateRatio(0.001d);
        ZipSecureFile.setMaxEntrySize(MAX_INPUT_BYTES);
        ZipSecureFile.setMaxTextSize(MAX_INPUT_BYTES);
    }
}
