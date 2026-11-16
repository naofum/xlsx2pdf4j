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

import org.junit.jupiter.api.Test;

import java.awt.Font;

import static org.assertj.core.api.Assertions.assertThat;

class FontResolverTest {

    @Test
    void fallsBackWhenLatinFontCannotDisplayJapanese() {
        FontResolver resolver = new FontResolver();
        String text = "ここの定義はXMLに依存";

        Font font = resolver.resolveForText("Calibri", 11, false, false, text);

        assertThat(font.canDisplayUpTo(text)).isEqualTo(-1);
        assertThat(font.getFamily()).isNotEqualTo("Calibri");
    }

    @Test
    void appliesFontScale() {
        FontResolver resolver = new FontResolver(0.90);

        Font font = resolver.resolve("MS PGothic", 10, false, false);

        assertThat(font.getSize2D()).isEqualTo(9.0f);
        assertThat(resolver.scaleSize(10)).isEqualTo(9.0);
    }

    @Test
    void keepsLatinFontForLatinOnlyText() {
        FontResolver resolver = new FontResolver();

        Font font = resolver.resolveForText("Calibri", 11, false, false, "XML");

        assertThat(font.canDisplayUpTo("XML")).isEqualTo(-1);
    }
}
