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
 * Fatal conversion errors.
 */
public enum ErrorCode {
    INPUT_NOT_FOUND,
    INPUT_FORMAT,
    INPUT_ENCRYPTED,
    INPUT_CORRUPT,
    OUTPUT_IO,
    LIMIT_EXCEEDED,
    FONT,
    UNSUPPORTED,
    INTERNAL
}
