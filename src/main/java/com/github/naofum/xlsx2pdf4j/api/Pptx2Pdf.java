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

package com.github.naofum.xlsx2pdf4j.api;

import com.github.naofum.xlsx2pdf4j.internal.pptx.PptxConversionService;

import java.nio.file.Path;

/**
 * Public entry point for PowerPoint (.pptx/.pptm) to PDF conversion.
 *
 * <p>The common {@link ConversionOptions} are used as follows: {@code dpi},
 * {@code textLayerMode}, {@code imageCompression}, {@code jpegQuality},
 * {@code maxPages}, {@code maxOutputBytes} and {@code strict} apply.
 * {@code sheets} selects slides by 1-based number or slide title. Excel-only
 * options such as {@code printScale} and {@code pageNumberingMode} are
 * ignored.</p>
 */
public final class Pptx2Pdf {

    private Pptx2Pdf() {
    }

    public static ConversionResult convert(Path input, Path output) throws Xlsx2PdfException {
        return convert(input, output, ConversionOptions.defaults());
    }

    public static ConversionResult convert(Path input, Path output,
                                           ConversionOptions options) throws Xlsx2PdfException {
        return new PptxConversionService().convert(input, output, options);
    }
}
