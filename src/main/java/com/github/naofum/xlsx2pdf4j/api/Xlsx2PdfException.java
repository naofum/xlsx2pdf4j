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
 * Fatal error raised by the converter.
 */
public class Xlsx2PdfException extends Exception {

    private final ErrorCode code;
    private final String sheetName;
    private final String location;

    public Xlsx2PdfException(ErrorCode code, String message) {
        this(code, message, null, null, null);
    }

    public Xlsx2PdfException(ErrorCode code, String message, Throwable cause) {
        this(code, message, null, null, cause);
    }

    public Xlsx2PdfException(ErrorCode code, String message, String sheetName, String location) {
        this(code, message, sheetName, location, null);
    }

    public Xlsx2PdfException(ErrorCode code, String message, String sheetName, String location, Throwable cause) {
        super(message, cause);
        this.code = code;
        this.sheetName = sheetName;
        this.location = location;
    }

    public ErrorCode code() {
        return code;
    }

    public String sheetName() {
        return sheetName;
    }

    public String location() {
        return location;
    }
}
