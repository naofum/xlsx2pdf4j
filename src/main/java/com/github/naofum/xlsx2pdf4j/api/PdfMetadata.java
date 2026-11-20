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

import java.time.Instant;

/**
 * Optional PDF document metadata.
 *
 * <p>Every field defaults to {@code null}. A null value keeps the converter's
 * existing behavior: {@code title} falls back to the input file name,
 * {@code creator} and {@code producer} fall back to "xlsx2pdf4j", and the
 * other attributes are left unset.
 */
public final class PdfMetadata {

    private final String title;
    private final String author;
    private final String subject;
    private final String keywords;
    private final String creator;
    private final String producer;
    private final Instant creationDate;
    private final Instant modificationDate;

    private PdfMetadata(Builder builder) {
        this.title = builder.title;
        this.author = builder.author;
        this.subject = builder.subject;
        this.keywords = builder.keywords;
        this.creator = builder.creator;
        this.producer = builder.producer;
        this.creationDate = builder.creationDate;
        this.modificationDate = builder.modificationDate;
    }

    public static Builder builder() {
        return new Builder();
    }

    public String title() {
        return title;
    }

    public String author() {
        return author;
    }

    public String subject() {
        return subject;
    }

    public String keywords() {
        return keywords;
    }

    public String creator() {
        return creator;
    }

    public String producer() {
        return producer;
    }

    /** Creation time. Converted to UTC when written to the PDF Info dictionary. */
    public Instant creationDate() {
        return creationDate;
    }

    /** Modification time. Converted to UTC when written to the PDF Info dictionary. */
    public Instant modificationDate() {
        return modificationDate;
    }

    public static final class Builder {

        private String title;
        private String author;
        private String subject;
        private String keywords;
        private String creator;
        private String producer;
        private Instant creationDate;
        private Instant modificationDate;

        private Builder() {
        }

        /** Null keeps the input file name as the PDF title. */
        public Builder title(String title) {
            this.title = title;
            return this;
        }

        public Builder author(String author) {
            this.author = author;
            return this;
        }

        public Builder subject(String subject) {
            this.subject = subject;
            return this;
        }

        public Builder keywords(String keywords) {
            this.keywords = keywords;
            return this;
        }

        /** Null keeps "xlsx2pdf4j" as the creator. */
        public Builder creator(String creator) {
            this.creator = creator;
            return this;
        }

        /** Null keeps "xlsx2pdf4j" as the producer. */
        public Builder producer(String producer) {
            this.producer = producer;
            return this;
        }

        public Builder creationDate(Instant creationDate) {
            this.creationDate = creationDate;
            return this;
        }

        public Builder modificationDate(Instant modificationDate) {
            this.modificationDate = modificationDate;
            return this;
        }

        public PdfMetadata build() {
            return new PdfMetadata(this);
        }
    }
}
