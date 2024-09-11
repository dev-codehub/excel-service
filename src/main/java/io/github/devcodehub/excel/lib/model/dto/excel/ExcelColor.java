package io.github.devcodehub.excel.lib.model.dto.excel;

import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * This enum represents the colors used in Excel files.
 * If the color already exists in the Excel palettes, please use the same name to maintain consistency.
 */
@Getter
@AllArgsConstructor
public enum ExcelColor implements ExcelColorBase {
    // Black
    BLACK(new XSSFColor(new byte[]{0, 0, 0})),

    // Blue
    BLUE(new XSSFColor(new byte[]{(byte) 0, (byte) 176, (byte) 240})),
    BLUE_ACCENT_1_LIGHTER_40(new XSSFColor(new byte[]{(byte) 142, (byte) 169, (byte) 219})),
    LIGHT_BLUE(new XSSFColor(new byte[]{(byte) 221, (byte) 235, (byte) 247})),

    // Brown
    BROWN(new XSSFColor(new byte[]{(byte) 153, (byte) 102, (byte) 0})),
    LIGHT_BROWN(new XSSFColor(new byte[]{(byte) 255, (byte) 234, (byte) 213})),

    // Cyan
    LIGHT_CYAN(new XSSFColor(new byte[]{(byte) 230, (byte) 255, (byte) 255})),

    // Green
    GREEN(new XSSFColor(new byte[]{(byte) 112, (byte) 173, (byte) 71})),
    LIGHT_GREEN(new XSSFColor(new byte[]{(byte) 226, (byte) 239, (byte) 218})),

    // Grey
    DARK_GREY(new XSSFColor(new byte[]{(byte) 150, (byte) 150, (byte) 150})),
    GREY(new XSSFColor(new byte[]{(byte) 237, (byte) 237, (byte) 237})),
    LIGHT_GREY(new XSSFColor(new byte[]{(byte) 212, (byte) 212, (byte) 212})),

    // Mint
    DARK_MINT(new XSSFColor(new byte[]{(byte) 73, (byte) 126, (byte) 118})),
    LIGHT_DARK_MINT(new XSSFColor(new byte[]{(byte) 219, (byte) 242, (byte) 237})),

    // Orange
    ORANGE(new XSSFColor(new byte[]{(byte) 237, (byte) 125, (byte) 49})),
    LIGHT_ORANGE(new XSSFColor(new byte[]{(byte) 252, (byte) 228, (byte) 214})),

    // Purple
    PURPLE(new XSSFColor(new byte[]{(byte) 191, (byte) 149, (byte) 223})),
    LIGHT_PURPLE(new XSSFColor(new byte[]{(byte) 239, (byte) 231, (byte) 255})),

    // Pink
    PINK(new XSSFColor(new byte[]{(byte) 255, (byte) 174, (byte) 201})),
    LIGHT_PINK(new XSSFColor(new byte[]{(byte) 255, (byte) 237, (byte) 243})),

    // Red
    RED(new XSSFColor(new byte[]{(byte) 192, (byte) 0, (byte) 0})),
    LIGHT_RED(new XSSFColor(new byte[]{(byte) 255, (byte) 225, (byte) 225})),

    // Teal
    TEAL(new XSSFColor(new byte[]{(byte) 91, (byte) 155, (byte) 213})),
    LIGHT_TEAL(new XSSFColor(new byte[]{(byte) 242, (byte) 248, (byte) 252})),

    // White
    WHITE(new XSSFColor(new byte[]{(byte) 255, (byte) 255, (byte) 255})),

    // Yellow
    YELLOW(new XSSFColor(new byte[]{(byte) 255, (byte) 255, (byte) 0})),
    LIGHT_YELLOW(new XSSFColor(new byte[]{(byte) 255, (byte) 255, (byte) 204})),

    // Yellow-brown
    YELLOW_BROWN(new XSSFColor(new byte[]{(byte) 255, (byte) 192, (byte) 0})),
    LIGHT_YELLOW_BROWN(new XSSFColor(new byte[]{(byte) 255, (byte) 242, (byte) 204}));

    final XSSFColor color;
}
