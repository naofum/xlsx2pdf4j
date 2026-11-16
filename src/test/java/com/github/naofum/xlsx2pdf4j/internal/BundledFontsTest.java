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

import static org.assertj.core.api.Assertions.assertThat;

class BundledFontsTest {

    @Test
    void loadsBundledJapaneseFonts() {
        assertThat(BundledFonts.isAvailable()).isTrue();
        assertThat(BundledFonts.sans()).isNotNull();
        assertThat(BundledFonts.serif()).isNotNull();
        assertThat(BundledFonts.sans().getFamily()).contains("Noto Sans JP");
        assertThat(BundledFonts.serif().getFamily()).contains("Noto Serif JP");
    }
}
