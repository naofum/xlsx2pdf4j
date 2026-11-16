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

import com.github.naofum.xlsx2pdf4j.internal.TextLayoutEngine;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.junit.jupiter.api.Test;

import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.Rectangle;
import java.awt.RenderingHints;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Regression tests for cell text rotation. Rotated text is laid out in a
 * height/width-swapped rectangle so that it is centered inside the unrotated
 * cell clip instead of starting at the original left edge and being clipped.
 */
class CellTextRotationTest {

    private static final Rectangle CELL = new Rectangle(50, 50, 200, 60);

    @Test
    void rotatedNinetyDegreesIsCenteredAndNotClipped() {
        BufferedImage image = render("ROT", 90);
        Rectangle ink = inkBounds(image);
        assertThat(ink).as("rotated text produces ink").isNotNull();

        assertThat(ink.width)
                .as("90 degree text is narrow")
                .isBetween(10, 40);
        assertThat(ink.height)
                .as("90 degree text keeps its full text length")
                .isBetween(35, 60);
        assertThat(ink.y).isGreaterThanOrEqualTo(CELL.y);
        assertThat(ink.y + ink.height).isLessThanOrEqualTo(CELL.y + CELL.height);
        assertThat(Math.abs(ink.y + ink.height / 2 - (CELL.y + CELL.height / 2)))
                .as("rotated text is vertically centered")
                .isLessThan(8);
    }

    @Test
    void rotatedFortyFiveDegreesIsCenteredAndNotClipped() {
        BufferedImage image = render("ROT", 45);
        Rectangle ink = inkBounds(image);
        assertThat(ink).as("rotated text produces ink").isNotNull();

        assertThat(ink.width)
                .as("45 degree text is diagonal")
                .isBetween(30, 55);
        assertThat(ink.height)
                .as("45 degree text is diagonal")
                .isBetween(30, 55);
        assertThat(ink.x).isGreaterThanOrEqualTo(CELL.x);
        assertThat(ink.x + ink.width).isLessThanOrEqualTo(CELL.x + CELL.width);
        assertThat(ink.y).isGreaterThanOrEqualTo(CELL.y);
        assertThat(ink.y + ink.height).isLessThanOrEqualTo(CELL.y + CELL.height);
    }

    private BufferedImage render(String text, int rotation) {
        BufferedImage image = new BufferedImage(300, 160, BufferedImage.TYPE_INT_RGB);
        Graphics2D graphics = image.createGraphics();
        try {
            graphics.setColor(Color.WHITE);
            graphics.fillRect(0, 0, image.getWidth(), image.getHeight());
            graphics.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING,
                    RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
            new TextLayoutEngine().drawText(graphics, text,
                    new Font(Font.SANS_SERIF, Font.BOLD, 24), Color.BLACK,
                    new Rectangle2D.Double(CELL.x, CELL.y, CELL.width, CELL.height),
                    HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
                    false, false, rotation, 0, false, false);
        } finally {
            graphics.dispose();
        }
        return image;
    }

    private Rectangle inkBounds(BufferedImage image) {
        int minX = Integer.MAX_VALUE;
        int minY = Integer.MAX_VALUE;
        int maxX = -1;
        int maxY = -1;
        for (int y = 0; y < image.getHeight(); y++) {
            for (int x = 0; x < image.getWidth(); x++) {
                int rgb = image.getRGB(x, y);
                int red = (rgb >> 16) & 0xFF;
                int green = (rgb >> 8) & 0xFF;
                int blue = rgb & 0xFF;
                if (red < 128 && green < 128 && blue < 128) {
                    minX = Math.min(minX, x);
                    minY = Math.min(minY, y);
                    maxX = Math.max(maxX, x);
                    maxY = Math.max(maxY, y);
                }
            }
        }
        return maxX < 0 ? null : new Rectangle(minX, minY, maxX - minX + 1, maxY - minY + 1);
    }
}
