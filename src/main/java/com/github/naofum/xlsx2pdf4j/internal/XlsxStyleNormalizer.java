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

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTXf;

/**
 * Applies Excel's interpretation of omitted style flags.
 *
 * <p>Files written by some generators (for example openpyxl) set a non-zero
 * borderId but omit applyBorder. Excel still renders the border, while POI
 * returns BorderStyle.NONE unless applyBorder is true. Normalizing the flag
 * lets the existing renderer see the border without changing the source file.
 */
final class XlsxStyleNormalizer {

    private XlsxStyleNormalizer() {
    }

    static void applyExcelDefaults(XSSFWorkbook workbook) {
        int count = workbook.getNumCellStyles();
        for (int index = 0; index < count; index++) {
            XSSFCellStyle style = workbook.getCellStyleAt(index);
            if (style == null) {
                continue;
            }
            CTXf xf = style.getCoreXf();
            if (xf != null && xf.getBorderId() != 0 && !xf.isSetApplyBorder()) {
                // Excel applies the referenced border when the attribute is
                // absent. An explicit applyBorder="0" is left untouched.
                xf.setApplyBorder(true);
            }
        }
    }
}
