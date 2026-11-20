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
import java.util.LinkedHashSet;
import java.util.Set;

/**
 * Immutable options for PowerPoint to PDF conversion.
 */
public final class PptxConversionOptions {

    public static final int DEFAULT_MAX_SLIDES = 200;
    public static final long DEFAULT_MAX_OUTPUT_BYTES = 50L * 1024 * 1024;

    private final Set<Integer> slides;
    private final int dpi;
    private final TextLayerMode textLayerMode;
    private final ImageCompression imageCompression;
    private final float jpegQuality;
    private final int maxSlides;
    private final long maxOutputBytes;
    private final boolean strict;
    private final PdfMetadata pdfMetadata;

    private PptxConversionOptions(Builder builder) {
        this.slides = Collections.unmodifiableSet(new LinkedHashSet<>(builder.slides));
        this.dpi = builder.dpi;
        this.textLayerMode = builder.textLayerMode;
        this.imageCompression = builder.imageCompression;
        this.jpegQuality = builder.jpegQuality;
        this.maxSlides = builder.maxSlides;
        this.maxOutputBytes = builder.maxOutputBytes;
        this.strict = builder.strict;
        this.pdfMetadata = builder.pdfMetadata;
    }

    public static Builder builder() {
        return new Builder();
    }

    public static PptxConversionOptions defaults() {
        return builder().build();
    }

    /** Empty means all visible slides. */
    public Set<Integer> slides() {
        return slides;
    }

    public int dpi() {
        return dpi;
    }

    public TextLayerMode textLayerMode() {
        return textLayerMode;
    }

    public ImageCompression imageCompression() {
        return imageCompression;
    }

    public float jpegQuality() {
        return jpegQuality;
    }

    public int maxSlides() {
        return maxSlides;
    }

    public long maxOutputBytes() {
        return maxOutputBytes;
    }

    public boolean strict() {
        return strict;
    }

    /** Optional PDF metadata. Null keeps the existing PDF attributes. */
    public PdfMetadata pdfMetadata() {
        return pdfMetadata;
    }

    public static final class Builder {

        private final Set<Integer> slides = new LinkedHashSet<>();
        private int dpi = 200;
        private TextLayerMode textLayerMode = TextLayerMode.INVISIBLE;
        private ImageCompression imageCompression = ImageCompression.AUTO;
        private float jpegQuality = 0.9f;
        private int maxSlides = DEFAULT_MAX_SLIDES;
        private long maxOutputBytes = DEFAULT_MAX_OUTPUT_BYTES;
        private boolean strict;
        private PdfMetadata pdfMetadata;

        private Builder() {
        }

        /** 1-based slide numbers. Empty selects all visible slides. */
        public Builder slides(int... slideNumbers) {
            if (slideNumbers != null) {
                for (int number : slideNumbers) {
                    slide(number);
                }
            }
            return this;
        }

        public Builder slide(int slideNumber) {
            if (slideNumber < 1) {
                throw new IllegalArgumentException("slideNumber must be positive: " + slideNumber);
            }
            slides.add(slideNumber);
            return this;
        }

        public Builder dpi(int dpi) {
            if (dpi < 72 || dpi > 600) {
                throw new IllegalArgumentException("dpi must be between 72 and 600: " + dpi);
            }
            this.dpi = dpi;
            return this;
        }

        public Builder textLayerMode(TextLayerMode mode) {
            this.textLayerMode = mode == null ? TextLayerMode.INVISIBLE : mode;
            return this;
        }

        public Builder imageCompression(ImageCompression compression) {
            this.imageCompression = compression == null ? ImageCompression.AUTO : compression;
            return this;
        }

        public Builder jpegQuality(float jpegQuality) {
            if (jpegQuality < 0.1f || jpegQuality > 1.0f) {
                throw new IllegalArgumentException("jpegQuality must be between 0.1 and 1.0");
            }
            this.jpegQuality = jpegQuality;
            return this;
        }

        public Builder maxSlides(int maxSlides) {
            if (maxSlides < 1) {
                throw new IllegalArgumentException("maxSlides must be positive");
            }
            this.maxSlides = maxSlides;
            return this;
        }

        public Builder maxOutputBytes(long maxOutputBytes) {
            if (maxOutputBytes < 1) {
                throw new IllegalArgumentException("maxOutputBytes must be positive");
            }
            this.maxOutputBytes = maxOutputBytes;
            return this;
        }

        public Builder strict(boolean strict) {
            this.strict = strict;
            return this;
        }

        /** Optional PDF metadata. Null keeps the existing attributes. */
        public Builder pdfMetadata(PdfMetadata metadata) {
            this.pdfMetadata = metadata;
            return this;
        }

        public PptxConversionOptions build() {
            return new PptxConversionOptions(this);
        }
    }
}
