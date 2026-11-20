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

import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

/**
 * Immutable conversion options.
 */
public final class ConversionOptions {

    public static final long DEFAULT_MAX_OUTPUT_BYTES = 50L * 1024 * 1024;
    public static final int DEFAULT_MAX_PAGES = 200;
    public static final double DEFAULT_FONT_SCALE = 1.0;

    private final Set<String> sheets;
    private final int dpi;
    private final boolean strict;
    private final int maxPages;
    private final long maxOutputBytes;
    private final Locale locale;
    private final ImageCompression imageCompression;
    private final float jpegQuality;
    private final TextLayerMode textLayerMode;
    private final double fontScale;
    private final double printScale;
    private final PageNumberingMode defaultPageNumberingMode;
    private final Map<String, PageNumberingMode> pageNumberingModes;
    private final PdfMetadata pdfMetadata;
    private final boolean sheetBookmarks;

    private ConversionOptions(Builder builder) {
        this.sheets = Collections.unmodifiableSet(new LinkedHashSet<>(builder.sheets));
        this.dpi = builder.dpi;
        this.strict = builder.strict;
        this.maxPages = builder.maxPages;
        this.maxOutputBytes = builder.maxOutputBytes;
        this.locale = builder.locale;
        this.imageCompression = builder.imageCompression;
        this.jpegQuality = builder.jpegQuality;
        this.textLayerMode = builder.textLayerMode;
        this.fontScale = builder.fontScale;
        this.printScale = builder.printScale;
        this.defaultPageNumberingMode = builder.defaultPageNumberingMode;
        this.pageNumberingModes = Collections.unmodifiableMap(new LinkedHashMap<>(builder.pageNumberingModes));
        this.pdfMetadata = builder.pdfMetadata;
        this.sheetBookmarks = builder.sheetBookmarks;
    }

    public static Builder builder() {
        return new Builder();
    }

    public static ConversionOptions defaults() {
        return builder().build();
    }

    /** Empty means all visible sheets. */
    public Set<String> sheets() {
        return sheets;
    }

    public int dpi() {
        return dpi;
    }

    public boolean strict() {
        return strict;
    }

    public int maxPages() {
        return maxPages;
    }

    public long maxOutputBytes() {
        return maxOutputBytes;
    }

    public Locale locale() {
        return locale;
    }

    public ImageCompression imageCompression() {
        return imageCompression;
    }

    public float jpegQuality() {
        return jpegQuality;
    }

    public TextLayerMode textLayerMode() {
        return textLayerMode;
    }

    /** Multiplier applied to every rendered font size. */
    public double fontScale() {
        return fontScale;
    }

    /**
     * Explicit print scale override. Zero means "use the scale stored in the
     * workbook".
     */
    public double printScale() {
        return printScale;
    }

    public PageNumberingMode pageNumberingMode(String sheetName) {
        return pageNumberingModes.getOrDefault(sheetName, defaultPageNumberingMode);
    }

    public PageNumberingMode defaultPageNumberingMode() {
        return defaultPageNumberingMode;
    }

    public Map<String, PageNumberingMode> pageNumberingModes() {
        return pageNumberingModes;
    }

    /** Optional PDF metadata. Null keeps the existing PDF attributes. */
    public PdfMetadata pdfMetadata() {
        return pdfMetadata;
    }

    /** Whether each worksheet gets a PDF bookmark at its first page. */
    public boolean sheetBookmarks() {
        return sheetBookmarks;
    }

    public static final class Builder {
        private final Set<String> sheets = new LinkedHashSet<>();
        private int dpi = 200;
        private boolean strict;
        private int maxPages = DEFAULT_MAX_PAGES;
        private long maxOutputBytes = DEFAULT_MAX_OUTPUT_BYTES;
        private Locale locale = Locale.JAPAN;
        private ImageCompression imageCompression = ImageCompression.AUTO;
        private float jpegQuality = 0.9f;
        private TextLayerMode textLayerMode = TextLayerMode.INVISIBLE;
        private double fontScale = DEFAULT_FONT_SCALE;
        private double printScale;
        private PageNumberingMode defaultPageNumberingMode = PageNumberingMode.EXCEL_AUTOMATIC;
        private final Map<String, PageNumberingMode> pageNumberingModes = new LinkedHashMap<>();
        private PdfMetadata pdfMetadata;
        private boolean sheetBookmarks;

        private Builder() {
        }

        public Builder sheets(String... names) {
            if (names != null) {
                Collections.addAll(sheets, names);
            }
            return this;
        }

        public Builder sheet(String name) {
            if (name != null) {
                sheets.add(name);
            }
            return this;
        }

        public Builder dpi(int dpi) {
            if (dpi < 72 || dpi > 600) {
                throw new IllegalArgumentException("dpi must be between 72 and 600: " + dpi);
            }
            this.dpi = dpi;
            return this;
        }

        public Builder strict(boolean strict) {
            this.strict = strict;
            return this;
        }

        public Builder maxPages(int maxPages) {
            if (maxPages < 1) {
                throw new IllegalArgumentException("maxPages must be positive");
            }
            this.maxPages = maxPages;
            return this;
        }

        public Builder maxOutputBytes(long maxOutputBytes) {
            if (maxOutputBytes < 1) {
                throw new IllegalArgumentException("maxOutputBytes must be positive");
            }
            this.maxOutputBytes = maxOutputBytes;
            return this;
        }

        public Builder locale(Locale locale) {
            this.locale = locale == null ? Locale.JAPAN : locale;
            return this;
        }

        public Builder imageCompression(ImageCompression imageCompression) {
            this.imageCompression = imageCompression == null ? ImageCompression.AUTO : imageCompression;
            return this;
        }

        public Builder jpegQuality(float jpegQuality) {
            if (jpegQuality < 0.1f || jpegQuality > 1.0f) {
                throw new IllegalArgumentException("jpegQuality must be between 0.1 and 1.0");
            }
            this.jpegQuality = jpegQuality;
            return this;
        }

        public Builder textLayerMode(TextLayerMode mode) {
            this.textLayerMode = mode == null ? TextLayerMode.INVISIBLE : mode;
            return this;
        }

        /**
         * Multiplier applied to all cell, shape, header and footer fonts.
         * Defaults to 1.0.
         */
        public Builder fontScale(double fontScale) {
            if (fontScale < 0.1 || fontScale > 4.0) {
                throw new IllegalArgumentException("fontScale must be between 0.1 and 4.0");
            }
            this.fontScale = fontScale;
            return this;
        }

        /**
         * Overrides Excel's print scale. Zero keeps the workbook setting.
         */
        public Builder printScale(double printScale) {
            if (printScale != 0 && (printScale < 0.1 || printScale > 4.0)) {
                throw new IllegalArgumentException("printScale must be 0 or between 0.1 and 4.0");
            }
            this.printScale = printScale;
            return this;
        }

        public Builder pageNumberingMode(PageNumberingMode mode) {
            this.defaultPageNumberingMode = mode == null ? PageNumberingMode.EXCEL_AUTOMATIC : mode;
            return this;
        }

        public Builder pageNumberingMode(String sheetName, PageNumberingMode mode) {
            if (sheetName == null || sheetName.isEmpty()) {
                throw new IllegalArgumentException("sheetName must not be empty");
            }
            this.pageNumberingModes.put(sheetName, mode == null ? PageNumberingMode.EXCEL_AUTOMATIC : mode);
            return this;
        }

        /** Optional PDF metadata. Null keeps the existing attributes. */
        public Builder pdfMetadata(PdfMetadata metadata) {
            this.pdfMetadata = metadata;
            return this;
        }

        /** Adds a PDF bookmark for each worksheet at its first page. */
        public Builder sheetBookmarks(boolean sheetBookmarks) {
            this.sheetBookmarks = sheetBookmarks;
            return this;
        }

        public ConversionOptions build() {
            return new ConversionOptions(this);
        }
    }
}
