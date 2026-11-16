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

import javax.imageio.ImageIO;
import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Locale;

/**
 * Writes a three-panel PNG for each golden-comparison page:
 * candidate, golden reference, and an amplified absolute difference.
 */
final class PageDiffWriter {

    private static final int GAP = 6;
    private static final int HEADER = 22;
    private static final int AMPLIFY = 4;

    private PageDiffWriter() {
    }

    static void write(Path directory, int pageNumber,
                      BufferedImage candidate, BufferedImage golden) throws IOException {
        BufferedImage reference = sameSize(golden,
                candidate.getWidth(), candidate.getHeight());
        BufferedImage difference = difference(candidate, reference);

        int width = candidate.getWidth();
        int height = candidate.getHeight();
        int second = width + GAP;
        int third = second + width + GAP;
        BufferedImage output = new BufferedImage(third + width, HEADER + height,
                BufferedImage.TYPE_INT_RGB);
        Graphics2D g = output.createGraphics();
        try {
            g.setColor(Color.WHITE);
            g.fillRect(0, 0, output.getWidth(), output.getHeight());
            g.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING,
                    RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
            g.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 12));
            g.setColor(new Color(0x333333));
            g.drawString("page " + pageNumber + " / candidate", 4, 15);
            g.drawString("golden", second + 4, 15);
            g.drawString("difference x" + AMPLIFY, third + 4, 15);

            g.drawImage(candidate, 0, HEADER, null);
            g.drawImage(reference, second, HEADER, null);
            g.drawImage(difference, third, HEADER, null);

            g.setColor(new Color(0xDDDDDD));
            g.drawRect(0, HEADER - 1, width - 1, height);
            g.drawRect(second, HEADER - 1, width - 1, height);
            g.drawRect(third, HEADER - 1, width - 1, height);
        } finally {
            g.dispose();
        }
        ImageIO.write(output, "png", directory.resolve(
                String.format(Locale.ROOT, "page-%02d-diff.png", pageNumber)).toFile());
    }

    static void writeSummary(Path directory, double[] differences) throws IOException {
        int pages = differences.length;
        double total = 0;
        for (double difference : differences) {
            total += difference;
        }
        double mean = pages == 0 ? 0 : total / pages;

        StringBuilder summary = new StringBuilder();
        summary.append("pages: ").append(pages).append('\n');
        summary.append(String.format(Locale.ROOT, "mean: %.6f%n", mean));
        summary.append("\npage,difference\n");
        for (int i = 0; i < pages; i++) {
            summary.append(i + 1).append(',')
                    .append(String.format(Locale.ROOT, "%.6f", differences[i]))
                    .append('\n');
        }

        int worstCount = Math.min(10, pages);
        boolean[] used = new boolean[pages];
        summary.append("\nworst pages:\n");
        for (int rank = 0; rank < worstCount; rank++) {
            int worst = -1;
            for (int i = 0; i < pages; i++) {
                if (!used[i] && (worst < 0 || differences[i] > differences[worst])) {
                    worst = i;
                }
            }
            if (worst < 0) {
                break;
            }
            used[worst] = true;
            summary.append(String.format(Locale.ROOT, "page %d: %.6f%n",
                    worst + 1, differences[worst]));
        }

        Files.writeString(directory.resolve("summary.txt"), summary.toString(),
                StandardCharsets.UTF_8);
    }

    private static BufferedImage difference(BufferedImage candidate, BufferedImage golden) {
        BufferedImage output = new BufferedImage(candidate.getWidth(), candidate.getHeight(),
                BufferedImage.TYPE_INT_RGB);
        for (int y = 0; y < candidate.getHeight(); y++) {
            for (int x = 0; x < candidate.getWidth(); x++) {
                int a = candidate.getRGB(x, y);
                int b = golden.getRGB(x, y);
                int delta = Math.max(
                        Math.abs(((a >> 16) & 0xFF) - ((b >> 16) & 0xFF)),
                        Math.max(
                                Math.abs(((a >> 8) & 0xFF) - ((b >> 8) & 0xFF)),
                                Math.abs((a & 0xFF) - (b & 0xFF))));
                int amplified = Math.min(255, delta * AMPLIFY);
                output.setRGB(x, y, (255 << 16)
                        | ((255 - amplified) << 8)
                        | (255 - amplified));
            }
        }
        return output;
    }

    private static BufferedImage sameSize(BufferedImage source, int width, int height) {
        if (source.getWidth() == width && source.getHeight() == height) {
            return source;
        }
        BufferedImage result = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g = result.createGraphics();
        try {
            g.setRenderingHint(RenderingHints.KEY_INTERPOLATION,
                    RenderingHints.VALUE_INTERPOLATION_BILINEAR);
            g.drawImage(source, 0, 0, width, height, null);
        } finally {
            g.dispose();
        }
        return result;
    }
}
