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

import com.github.naofum.xlsx2pdf4j.api.ConversionOptions;
import com.github.naofum.xlsx2pdf4j.api.ConversionResult;
import com.github.naofum.xlsx2pdf4j.api.ConversionWarning;
import com.github.naofum.xlsx2pdf4j.api.ErrorCode;
import com.github.naofum.xlsx2pdf4j.api.Xlsx2PdfException;
import com.github.naofum.xlsx2pdf4j.api.ImageCompression;
import com.github.naofum.xlsx2pdf4j.api.PageNumberingMode;
import com.github.naofum.xlsx2pdf4j.api.WarningCode;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.image.BufferedImage;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Orchestrates reading, layout, rendering and PDF writing.
 */
public final class ConversionService {

    private static final Logger LOG = LoggerFactory.getLogger(ConversionService.class);
    private static final long MAX_INPUT_BYTES = 100L * 1024 * 1024;

    public ConversionResult convert(Path input, Path output, ConversionOptions options) throws Xlsx2PdfException {
        Instant started = Instant.now();
        validateInput(input);
        preparePoiSecurity();

        List<ConversionWarning> warnings = new ArrayList<>();
        Path absoluteOutput = output.toAbsolutePath();
        try (XSSFWorkbook workbook = openWorkbook(input)) {
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            DataFormatter formatter = new DataFormatter(options.locale());
            FontResolver fonts = new FontResolver(options.fontScale());
            TextLayoutEngine textEngine = new TextLayoutEngine();
            ShapeRenderer shapeRenderer = new ShapeRenderer(fonts, textEngine);
            HeaderFooterRenderer headerFooter = new HeaderFooterRenderer(fonts);
            LayoutEngine layoutEngine = new LayoutEngine();

            double columnWidthScale = columnWidthScale(workbook);
            List<SheetLayout> sheets = selectSheets(workbook, options, warnings, columnWidthScale);
            if (sheets.isEmpty()) {
                throw new Xlsx2PdfException(ErrorCode.INPUT_FORMAT, "no printable sheets were selected");
            }

            List<PageLayout> allPages = new ArrayList<>();
            Map<String, Integer> pagesBySheet = new LinkedHashMap<>();
            int nextPage = 1;
            for (SheetLayout sheet : sheets) {
                int start = nextPage;
                PageNumberingMode mode = options.pageNumberingMode(sheet.name);
                boolean useFirstPageNumber = mode == PageNumberingMode.USER_FIRST_PAGE
                        ? sheet.print.firstPageNumber > 0
                        : sheet.print.useFirstPageNumber && sheet.print.firstPageNumber > 0;
                if (useFirstPageNumber) {
                    start = sheet.print.firstPageNumber;
                }
                List<PageLayout> pages = layoutEngine.layout(sheet, start, options.printScale());
                allPages.addAll(pages);
                pagesBySheet.put(sheet.name, pages.size());
                nextPage = pages.isEmpty() ? nextPage : pages.get(pages.size() - 1).pageNumber + 1;
            }

            if (allPages.size() > options.maxPages()) {
                throw new Xlsx2PdfException(ErrorCode.LIMIT_EXCEEDED,
                        "page count " + allPages.size() + " exceeds limit " + options.maxPages());
            }

            ImageCompression compression = effectiveImageCompression(options);
            try (PdfWriter pdf = new PdfWriter(absoluteOutput, options.maxOutputBytes(),
                    compression, options.jpegQuality())) {
                int totalPages = allPages.size();
                String fileName = input.getFileName() == null ? "" : input.getFileName().toString();
                for (SheetLayout sheet : sheets) {
                    CellRenderer cellRenderer = new CellRenderer(workbook, sheet, fonts, textEngine,
                            formatter, evaluator, warnings);
                    PageRenderer pageRenderer = new PageRenderer(cellRenderer, shapeRenderer,
                            headerFooter, options.dpi());
                    for (PageLayout page : allPages) {
                        if (page.sheet != sheet) {
                            continue;
                        }
                        List<PlacedText> textLayer = options.textLayerMode()
                                == com.github.naofum.xlsx2pdf4j.api.TextLayerMode.INVISIBLE
                                ? new ArrayList<>()
                                : null;
                        BufferedImage image = pageRenderer.render(page, totalPages, fileName, warnings, textLayer);
                        checkStrict(options, warnings);
                        pdf.addPage(image, page.pageWidth, page.pageHeight, textLayer);
                        image.flush();
                    }
                }
                pdf.finish(fileName);
            }

            long size = Files.size(absoluteOutput);
            LOG.info("converted {} sheets, {} pages, {} bytes", sheets.size(), allPages.size(), size);
            return new ConversionResult(absoluteOutput, allPages.size(), pagesBySheet, warnings,
                    Duration.between(started, Instant.now()), size);
        } catch (Xlsx2PdfException e) {
            throw e;
        } catch (EncryptedDocumentException e) {
            throw new Xlsx2PdfException(ErrorCode.INPUT_ENCRYPTED, "the workbook is encrypted", e);
        } catch (InvalidFormatException e) {
            throw new Xlsx2PdfException(ErrorCode.INPUT_FORMAT, "invalid .xlsx package", e);
        } catch (IOException e) {
            throw new Xlsx2PdfException(ErrorCode.INPUT_CORRUPT, "cannot read workbook", e);
        } catch (RuntimeException e) {
            throw new Xlsx2PdfException(ErrorCode.INTERNAL, "unexpected conversion error", e);
        }
    }

    /**
     * In strict mode a critical warning aborts before the page is written and
     * before the temporary PDF is moved to the caller's path. Warnings that
     * only describe approximations (SHAPE_APPROXIMATED, PAGE_LIMIT, ...) keep
     * the conversion going.
     */
    private void checkStrict(ConversionOptions options, List<ConversionWarning> warnings)
            throws Xlsx2PdfException {
        if (!options.strict()) {
            return;
        }
        for (ConversionWarning warning : warnings) {
            if (warning.code() == WarningCode.SHAPE_UNSUPPORTED
                    || warning.code() == WarningCode.IMAGE_UNSUPPORTED
                    || warning.code() == WarningCode.FONT_MISSING
                    || warning.code() == WarningCode.FORMULA_NO_CACHE) {
                throw new Xlsx2PdfException(ErrorCode.UNSUPPORTED,
                        "strict mode: " + warning.message(), warning.sheetName(), warning.location());
            }
        }
    }

    private void validateInput(Path input) throws Xlsx2PdfException {
        if (input == null) {
            throw new Xlsx2PdfException(ErrorCode.INPUT_NOT_FOUND, "input path is null");
        }
        if (!Files.isRegularFile(input)) {
            throw new Xlsx2PdfException(ErrorCode.INPUT_NOT_FOUND, "input file not found: " + input);
        }
        String name = input.getFileName().toString().toLowerCase();
        if (!name.endsWith(".xlsx") && !name.endsWith(".xlsm")) {
            throw new Xlsx2PdfException(ErrorCode.INPUT_FORMAT, "only .xlsx/.xlsm is supported: " + input);
        }
        try {
            if (Files.size(input) > MAX_INPUT_BYTES) {
                throw new Xlsx2PdfException(ErrorCode.LIMIT_EXCEEDED,
                        "input file is larger than " + MAX_INPUT_BYTES + " bytes");
            }
        } catch (IOException e) {
            throw new Xlsx2PdfException(ErrorCode.INPUT_NOT_FOUND, "cannot stat input file", e);
        }
    }

    /**
     * PNG predictor encoding dominates 300 DPI conversion time. AUTO therefore
     * uses JPEG for high-resolution pages; callers can still force PNG.
     */
    private ImageCompression effectiveImageCompression(ConversionOptions options) {
        if (options.imageCompression() == ImageCompression.AUTO && options.dpi() >= 300) {
            return ImageCompression.JPEG;
        }
        return options.imageCompression();
    }

    /**
     * POI's column-width pixel conversion assumes a 7px default digit width.
     * Excel uses the metrics of the workbook's default font. Some workbooks
     * need a small calibration; the factor is kept conservative and can be
     * replaced by an exact MDW calculation later.
     */
    private double columnWidthScale(XSSFWorkbook workbook) {
        try {
            var excelFont = workbook.getFontAt(0);
            FontResolver baseFonts = new FontResolver(1.0);
            java.awt.Font font = baseFonts.resolve(excelFont.getFontName(),
                    excelFont.getFontHeightInPoints(), excelFont.getBold(), excelFont.getItalic());
            java.awt.font.FontRenderContext frc =
                    new java.awt.font.FontRenderContext(null, true, true);
            double mdwPixels = 0;
            for (char digit = '0'; digit <= '9'; digit++) {
                mdwPixels = Math.max(mdwPixels,
                        font.getStringBounds(String.valueOf(digit), frc).getWidth() * 96d / 72d);
            }
            if (mdwPixels <= 0) {
                return 1.0;
            }
            // POI's column-width pixel conversion assumes a 7px maximum digit
            // width. Scale it to the workbook default font's actual MDW.
            return Math.max(0.75, Math.min(1.10, mdwPixels / 7d));
        } catch (RuntimeException e) {
            return 1.0;
        }
    }

    private void preparePoiSecurity() {
        ZipSecureFile.setMinInflateRatio(0.001d);
        ZipSecureFile.setMaxEntrySize(MAX_INPUT_BYTES);
        ZipSecureFile.setMaxTextSize(MAX_INPUT_BYTES);
    }

    private XSSFWorkbook openWorkbook(Path input) throws IOException, InvalidFormatException, Xlsx2PdfException {
        OPCPackage pkg = OPCPackage.open(input.toFile(), PackageAccess.READ);
        try {
            return new XSSFWorkbook(pkg);
        } catch (EncryptedDocumentException e) {
            pkg.close();
            throw e;
        } catch (RuntimeException e) {
            pkg.close();
            throw new Xlsx2PdfException(ErrorCode.INPUT_CORRUPT, "cannot parse workbook", e);
        }
    }

    private List<SheetLayout> selectSheets(XSSFWorkbook workbook, ConversionOptions options,
                                           List<ConversionWarning> warnings, double columnWidthScale) {
        Set<String> requested = options.sheets();
        List<SheetLayout> result = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            XSSFSheet sheet = workbook.getSheetAt(i);
            boolean visible = workbook.getSheetVisibility(i) == org.apache.poi.ss.usermodel.SheetVisibility.VISIBLE;
            if (!requested.isEmpty()) {
                if (!requested.contains(sheet.getSheetName())) {
                    continue;
                }
            } else if (!visible) {
                continue;
            }
            if (sheet.getPhysicalNumberOfRows() == 0 && sheet.getDrawingPatriarch() == null) {
                warnings.add(new ConversionWarning(WarningCode.PAGE_LIMIT, sheet.getSheetName(), "",
                        "sheet has no printable content"));
                continue;
            }
            result.add(SheetLayout.read(workbook, i, sheet, columnWidthScale));
        }
        return result;
    }
}
