package com.woter.fact.excel.util;

import org.apache.commons.lang3.time.DateFormatUtils;

import java.util.Date;

/**
 * @author woter
 * @date 2017-9-30 上午10:12:47
 */
public class DateUtil {

    public final static String TIME_FORMAT = "yyyy-MM-dd HH:mm:ss";

    public final static String DATE_FORMAT = "yyyy-MM-dd";

    public static String date2String(Date date, String pattern) {
        if (null == date) {
            return null;
        }
        return DateFormatUtils.format(date, pattern);
    }
}
