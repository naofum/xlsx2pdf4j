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

/**
 * Fatal error raised by the PowerPoint converter.
 *
 * <p>{@link #slideName()} is the "Slide N" label for the failing slide. The
 * inherited {@link #sheetName()} is kept for compatibility with the common
 * warning/error model.</p>
 */
public class Pptx2PdfException extends Xlsx2PdfException {

    public Pptx2PdfException(ErrorCode code, String message) {
        super(code, message);
    }

    public Pptx2PdfException(ErrorCode code, String message, Throwable cause) {
        super(code, message, cause);
    }

    public Pptx2PdfException(ErrorCode code, String message, String slideName, String location) {
        super(code, message, slideName, location);
    }

    public Pptx2PdfException(ErrorCode code, String message, String slideName,
                             String location, Throwable cause) {
        super(code, message, slideName, location, cause);
    }

    public String slideName() {
        return sheetName();
    }
}
