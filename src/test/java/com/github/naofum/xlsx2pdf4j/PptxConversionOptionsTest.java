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
package com.github.naofum.xlsx2pdf4j;

import com.github.naofum.xlsx2pdf4j.api.ImageCompression;
import com.github.naofum.xlsx2pdf4j.api.PptxConversionOptions;
import com.github.naofum.xlsx2pdf4j.api.TextLayerMode;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class PptxConversionOptionsTest {

    @Test
    void defaultsMatchTheDocumentedValues() {
        PptxConversionOptions options = PptxConversionOptions.defaults();

        assertThat(options.slides()).isEmpty();
        assertThat(options.dpi()).isEqualTo(200);
        assertThat(options.textLayerMode()).isEqualTo(TextLayerMode.INVISIBLE);
        assertThat(options.imageCompression()).isEqualTo(ImageCompression.AUTO);
        assertThat(options.jpegQuality()).isEqualTo(0.9f);
        assertThat(options.maxSlides()).isEqualTo(200);
        assertThat(options.maxOutputBytes()).isEqualTo(50L * 1024 * 1024);
        assertThat(options.strict()).isFalse();
    }

    @Test
    void builderAcceptsPptxSpecificValues() {
        PptxConversionOptions options = PptxConversionOptions.builder()
                .slides(2, 1)
                .slide(3)
                .dpi(300)
                .textLayerMode(TextLayerMode.NONE)
                .imageCompression(ImageCompression.JPEG)
                .jpegQuality(0.75f)
                .maxSlides(10)
                .maxOutputBytes(1024 * 1024)
                .strict(true)
                .build();

        assertThat(options.slides()).containsExactlyInAnyOrder(1, 2, 3);
        assertThat(options.dpi()).isEqualTo(300);
        assertThat(options.textLayerMode()).isEqualTo(TextLayerMode.NONE);
        assertThat(options.imageCompression()).isEqualTo(ImageCompression.JPEG);
        assertThat(options.jpegQuality()).isEqualTo(0.75f);
        assertThat(options.maxSlides()).isEqualTo(10);
        assertThat(options.maxOutputBytes()).isEqualTo(1024 * 1024);
        assertThat(options.strict()).isTrue();
    }

    @Test
    void rejectsInvalidValues() {
        assertThatThrownBy(() -> PptxConversionOptions.builder().dpi(71))
                .isInstanceOf(IllegalArgumentException.class);
        assertThatThrownBy(() -> PptxConversionOptions.builder().slide(0))
                .isInstanceOf(IllegalArgumentException.class);
        assertThatThrownBy(() -> PptxConversionOptions.builder().jpegQuality(0.05f))
                .isInstanceOf(IllegalArgumentException.class);
        assertThatThrownBy(() -> PptxConversionOptions.builder().maxSlides(0))
                .isInstanceOf(IllegalArgumentException.class);
        assertThatThrownBy(() -> PptxConversionOptions.builder().maxOutputBytes(0))
                .isInstanceOf(IllegalArgumentException.class);
    }
}
