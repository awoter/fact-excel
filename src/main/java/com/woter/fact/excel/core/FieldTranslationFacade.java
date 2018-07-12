package com.woter.fact.excel.core;

import com.woter.fact.excel.util.DateUtil;

/**
 * @author woter
 * @date 2017-9-30 上午10:28:41
 */
public final class FieldTranslationFacade {

    public static <T> DefaultFieldTranslation<T> defaultFieldTranslation(final String field) {
        return new DefaultFieldTranslation<T>(field);
    }

    public static <T> DefaultFieldTranslation<T> dateFieldTranslation(final String field) {
        return new DateFieldTranslation<T>(field, DateUtil.DATE_FORMAT);
    }

    public static <T> DefaultFieldTranslation<T> timeFieldTranslation(final String field) {
        return new DateFieldTranslation<T>(field, DateUtil.TIME_FORMAT);
    }

}
