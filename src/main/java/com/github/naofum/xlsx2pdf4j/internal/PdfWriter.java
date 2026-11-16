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

import com.github.naofum.xlsx2pdf4j.api.Xlsx2PdfException;
import com.github.naofum.xlsx2pdf4j.api.ErrorCode;
import com.github.naofum.xlsx2pdf4j.api.ImageCompression;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentInformation;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.graphics.image.JPEGFactory;
import org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.pdmodel.graphics.state.RenderingMode;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Locale;

/**
 * Writes rendered pages into a PDF. Pages are embedded as raster images because
 * the product does not require searchable text.
 */
public final class PdfWriter implements AutoCloseable {

    private final PDDocument document = new PDDocument();
    private final Path output;
    private final long maxOutputBytes;
    private final ImageCompression compression;
    private final float jpegQuality;
    private final Path temp;
    private PDType0Font textFont;
    private boolean textFontTried;

    public PdfWriter(Path output, long maxOutputBytes, ImageCompression compression, float jpegQuality)
            throws Xlsx2PdfException {
        this.output = output;
        this.maxOutputBytes = maxOutputBytes;
        this.compression = compression;
        this.jpegQuality = jpegQuality;
        try {
            Path parent = output.toAbsolutePath().getParent();
            if (parent != null) {
                Files.createDirectories(parent);
            }
            this.temp = Files.createTempFile(parent, "xlsx2pdf4j-", ".pdf");
        } catch (IOException e) {
            throw new Xlsx2PdfException(ErrorCode.OUTPUT_IO, "cannot create temporary PDF", e);
        }
    }

    public void addPage(BufferedImage image, double pageWidth, double pageHeight,
                        java.util.List<PlacedText> textLayer) throws Xlsx2PdfException {
        try {
            PDPage page = new PDPage(new PDRectangle((float) pageWidth, (float) pageHeight));
            document.addPage(page);
            PDImageXObject pdImage = switch (compression) {
                case JPEG -> JPEGFactory.createFromImage(document, image, jpegQuality);
                case PNG -> LosslessFactory.createFromImage(document, image);
                case AUTO -> LosslessFactory.createFromImage(document, image);
            };
            try (PDPageContentStream content = new PDPageContentStream(document, page)) {
                content.drawImage(pdImage, 0, 0, (float) pageWidth, (float) pageHeight);
                if (textLayer != null && !textLayer.isEmpty()) {
                    writeTextLayer(content, pageHeight, textLayer);
                }
            }
        } catch (IOException e) {
            throw new Xlsx2PdfException(ErrorCode.INTERNAL, "cannot add PDF page", e);
        }
    }

    private void writeTextLayer(PDPageContentStream content, double pageHeight,
                                java.util.List<PlacedText> textLayer) throws IOException {
        PDType0Font font = textFont();
        if (font == null) {
            return;
        }
        content.setRenderingMode(RenderingMode.NEITHER);
        for (PlacedText placed : textLayer) {
            String text = sanitize(placed.text());
            if (text.isEmpty()) {
                continue;
            }
            try {
                content.beginText();
                content.setFont(font, (float) Math.max(1, placed.fontSize()));
                content.newLineAtOffset((float) placed.x(), (float) (pageHeight - placed.baselineY()));
                content.showText(text);
                content.endText();
            } catch (RuntimeException | IOException e) {
                // Unsupported glyphs should not break the visual page.
                try {
                    content.endText();
                } catch (IOException ignored) {
                    // Best effort.
                }
            }
        }
        content.setRenderingMode(RenderingMode.FILL);
    }

    private PDType0Font textFont() throws IOException {
        if (textFontTried) {
            return textFont;
        }
        textFontTried = true;
        byte[] bytes = BundledFonts.sansBytes();
        if (bytes == null) {
            return null;
        }
        textFont = PDType0Font.load(document, new ByteArrayInputStream(bytes), true);
        return textFont;
    }

    private String sanitize(String value) {
        if (value == null) {
            return "";
        }
        StringBuilder out = new StringBuilder(value.length());
        for (int i = 0; i < value.length(); i++) {
            char c = value.charAt(i);
            if (c == '\r' || c == '\n' || c == '\t') {
                out.append(' ');
            } else if (!Character.isISOControl(c)) {
                out.append(c);
            }
        }
        return out.toString().trim();
    }

    public void finish(String title) throws Xlsx2PdfException {
        try {
            PDDocumentInformation info = document.getDocumentInformation();
            if (title != null && !title.isBlank()) {
                info.setTitle(title);
            }
            info.setProducer("xlsx2pdf4j");
            info.setCreator("xlsx2pdf4j");
            document.save(temp.toFile());
            document.close();
            long size = Files.size(temp);
            if (size > maxOutputBytes) {
                throw new Xlsx2PdfException(ErrorCode.LIMIT_EXCEEDED,
                        String.format(Locale.ROOT, "output PDF is %,d bytes; limit is %,d bytes", size, maxOutputBytes));
            }
            try {
                Files.move(temp, output, StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE);
            } catch (IOException atomicFailure) {
                Files.move(temp, output, StandardCopyOption.REPLACE_EXISTING);
            }
        } catch (Xlsx2PdfException e) {
            throw e;
        } catch (IOException e) {
            throw new Xlsx2PdfException(ErrorCode.OUTPUT_IO, "cannot write output PDF", e);
        } finally {
            try {
                document.close();
            } catch (IOException ignored) {
                // Already closed or partially written.
            }
            if (!Files.exists(output)) {
                try {
                    Files.deleteIfExists(temp);
                } catch (IOException ignored) {
                    // Best effort.
                }
            }
        }
    }

    @Override
    public void close() {
        try {
            document.close();
        } catch (IOException ignored) {
            // Best effort.
        }
        try {
            Files.deleteIfExists(temp);
        } catch (IOException ignored) {
            // Best effort.
        }
    }
}
