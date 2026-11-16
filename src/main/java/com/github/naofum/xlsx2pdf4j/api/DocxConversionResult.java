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

import java.nio.file.Path;
import java.time.Duration;
import java.util.List;

/**
 * Result of a successful Word to PDF conversion.
 */
public final class DocxConversionResult {

    private final Path output;
    private final int pageCount;
    private final List<ConversionWarning> warnings;
    private final Duration duration;
    private final long outputSize;

    public DocxConversionResult(Path output, int pageCount,
                                List<ConversionWarning> warnings,
                                Duration duration, long outputSize) {
        this.output = output;
        this.pageCount = pageCount;
        this.warnings = List.copyOf(warnings);
        this.duration = duration;
        this.outputSize = outputSize;
    }

    public Path output() {
        return output;
    }

    public int pageCount() {
        return pageCount;
    }

    public List<ConversionWarning> warnings() {
        return warnings;
    }

    public Duration duration() {
        return duration;
    }

    public long outputSize() {
        return outputSize;
    }
}
