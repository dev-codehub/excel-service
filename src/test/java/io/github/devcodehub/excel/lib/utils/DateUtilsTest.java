package io.github.devcodehub.excel.lib.utils;

import org.junit.jupiter.api.Test;

import java.util.Calendar;
import java.util.Date;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

class DateUtilsTest {

    @Test
    void getDateAsString_nullDate_returnsNull() {
        assertNull(DateUtils.getDateAsString(null));
    }

    @Test
    void getDateAsString_nullDate_withCustomPattern_returnsNull() {
        assertNull(DateUtils.getDateAsString(null, "dd-MM-yyyy"));
    }

    @Test
    void getDateAsString_defaultFormat_producesCorrectString() {
        Date date = buildDate(2024, Calendar.JANUARY, 15);

        assertEquals("2024/01/15", DateUtils.getDateAsString(date));
    }

    @Test
    void getDateAsString_customFormat_producesCorrectString() {
        Date date = buildDate(2024, Calendar.MARCH, 5);

        assertEquals("05-03-2024", DateUtils.getDateAsString(date, "dd-MM-yyyy"));
    }

    @Test
    void defaultDateFormat_constant_matchesExpectedPattern() {
        assertEquals("yyyy/MM/dd", DateUtils.DEFAULT_DATE_FORMAT);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static Date buildDate(int year, int month, int day) {
        Calendar cal = Calendar.getInstance();
        cal.set(year, month, day, 0, 0, 0);
        cal.set(Calendar.MILLISECOND, 0);
        return cal.getTime();
    }
}
