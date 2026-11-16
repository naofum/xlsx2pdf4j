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
 * How page numbers in headers/footers are calculated.
 */
public enum PageNumberingMode {
    /**
     * Follow Excel's page setup. A custom first page number is used only when
     * the workbook enables "use first page number".
     */
    EXCEL_AUTOMATIC,

    /**
     * Use the sheet's firstPageNumber when it is present and not the automatic
     * sentinel (4294967294), regardless of the useFirstPageNumber flag.
     */
    USER_FIRST_PAGE
}
