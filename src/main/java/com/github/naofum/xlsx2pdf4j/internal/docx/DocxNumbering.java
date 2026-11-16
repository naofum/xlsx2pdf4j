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

import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFNum;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;

import java.math.BigInteger;
import java.util.HashMap;
import java.util.Map;

/**
 * Minimal WordprocessingML numbering state. It handles the common decimal and
 * bullet formats used by headings and lists.
 */
final class DocxNumbering {

    record NumberedText(String text, double leftIndent, double hangingIndent) {
    }

    private final XWPFNumbering numbering;
    private final Map<Integer, int[]> counters = new HashMap<>();

    DocxNumbering(XWPFNumbering numbering) {
        this.numbering = numbering;
    }

    NumberedText next(int numId, int level) {
        if (numbering == null) {
            return null;
        }
        XWPFNum num = numbering.getNum(BigInteger.valueOf(numId));
        if (num == null || num.getCTNum() == null || num.getCTNum().getAbstractNumId() == null) {
            return null;
        }
        BigInteger abstractId = num.getCTNum().getAbstractNumId().getVal();
        XWPFAbstractNum abstractNum = numbering.getAbstractNum(abstractId);
        if (abstractNum == null || abstractNum.getCTAbstractNum() == null) {
            return null;
        }
        int safeLevel = Math.max(0, level);
        if (safeLevel >= abstractNum.getCTAbstractNum().sizeOfLvlArray()) {
            safeLevel = 0;
        }
        CTLvl lvl = abstractNum.getCTAbstractNum().getLvlArray(safeLevel);
        int[] values = counters.computeIfAbsent(numId, ignored -> new int[9]);
        if (values[safeLevel] <= 0) {
            int start = lvl.getStart() == null || lvl.getStart().getVal() == null
                    ? 1 : lvl.getStart().getVal().intValue();
            values[safeLevel] = Math.max(1, start);
        } else {
            values[safeLevel]++;
        }
        for (int i = safeLevel + 1; i < values.length; i++) {
            values[i] = 0;
        }
        String format = lvl.getNumFmt() == null ? "decimal"
                : String.valueOf(lvl.getNumFmt().getVal());
        String text = lvl.getLvlText() == null || lvl.getLvlText().getVal() == null
                ? "" : lvl.getLvlText().getVal();
        if ("bullet".equals(format) || "none".equals(format)) {
            return new NumberedText(text.isEmpty() ? "•" : text,
                    indent(lvl, true), indent(lvl, false));
        }
        for (int i = 0; i < values.length; i++) {
            int value = i == safeLevel ? values[i]
                    : (i < safeLevel ? values[i] : 0);
            text = text.replace("%" + (i + 1), value == 0 ? "" : format(value, format));
        }
        return new NumberedText(text, indent(lvl, true), indent(lvl, false));
    }

    private static double indent(CTLvl lvl, boolean left) {
        if (lvl == null || lvl.getPPr() == null || lvl.getPPr().getInd() == null) {
            return 0;
        }
        CTInd ind = lvl.getPPr().getInd();
        Object value = left ? ind.getLeft() : ind.getHanging();
        if (value == null) {
            return 0;
        }
        try {
            return new java.math.BigDecimal(value.toString()).doubleValue() / 20.0;
        } catch (NumberFormatException ignored) {
            return 0;
        }
    }

    private static String format(int value, String format) {
        return switch (format) {
            case "decimalFullWidth" -> fullWidth(Integer.toString(value));
            case "decimalEnclosedCircle" -> circled(value);
            case "aiueoFullWidth" -> aiueo(value);
            case "lowerLetter" -> String.valueOf((char) ('a' + (value - 1) % 26));
            case "upperLetter" -> String.valueOf((char) ('A' + (value - 1) % 26));
            case "lowerRoman" -> roman(value).toLowerCase(java.util.Locale.ROOT);
            case "upperRoman" -> roman(value);
            default -> Integer.toString(value);
        };
    }

    private static String fullWidth(String value) {
        StringBuilder result = new StringBuilder();
        for (char c : value.toCharArray()) {
            result.append(c >= '0' && c <= '9' ? (char) ('０' + (c - '0')) : c);
        }
        return result.toString();
    }

    private static String circled(int value) {
        if (value >= 1 && value <= 20) {
            return String.valueOf((char) ('①' + value - 1));
        }
        return "(" + value + ")";
    }

    private static String aiueo(int value) {
        String letters = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん";
        return value >= 1 && value <= letters.length()
                ? String.valueOf(letters.charAt(value - 1)) : Integer.toString(value);
    }

    private static String roman(int value) {
        int[] values = {1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1};
        String[] symbols = {"M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"};
        StringBuilder result = new StringBuilder();
        int remaining = value;
        for (int i = 0; i < values.length; i++) {
            while (remaining >= values[i]) {
                result.append(symbols[i]);
                remaining -= values[i];
            }
        }
        return result.toString();
    }
}
