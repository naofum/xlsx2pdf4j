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

import com.github.naofum.xlsx2pdf4j.api.ConversionOptions;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class ConversionOptionsTest {

    @Test
    void defaultsUseNinetyPercentFontsAndWorkbookPrintScale() {
        ConversionOptions options = ConversionOptions.defaults();

        assertThat(options.fontScale()).isEqualTo(1.0);
        assertThat(options.printScale()).isZero();
    }

    @Test
    void acceptsExplicitScales() {
        ConversionOptions options = ConversionOptions.builder()
                .fontScale(0.85)
                .printScale(0.75)
                .build();

        assertThat(options.fontScale()).isEqualTo(0.85);
        assertThat(options.printScale()).isEqualTo(0.75);
    }

    @Test
    void rejectsOutOfRangeScales() {
        assertThatThrownBy(() -> ConversionOptions.builder().fontScale(0.05))
                .isInstanceOf(IllegalArgumentException.class);
        assertThatThrownBy(() -> ConversionOptions.builder().printScale(5.0))
                .isInstanceOf(IllegalArgumentException.class);
    }
}
