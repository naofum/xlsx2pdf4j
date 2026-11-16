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

import com.github.naofum.xlsx2pdf4j.internal.docx.DocxConversionService;

import java.nio.file.Path;

/**
 * Public entry point for Word (.docx) to PDF conversion.
 */
public final class Docx2Pdf {

    private Docx2Pdf() {
    }

    public static DocxConversionResult convert(Path input, Path output)
            throws Docx2PdfException {
        return convert(input, output, DocxConversionOptions.defaults());
    }

    public static DocxConversionResult convert(Path input, Path output,
                                               DocxConversionOptions options)
            throws Docx2PdfException {
        return new DocxConversionService().convert(input, output, options);
    }
}
