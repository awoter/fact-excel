package com.woter.fact.excel.core;

import com.woter.fact.excel.util.DateUtil;
import org.apache.commons.lang3.StringUtils;

import java.util.Date;

/**
 * @author woter
 * @date 2017-9-30 上午10:08:52
 */
public class DateFieldTranslation<T> extends DefaultFieldTranslation<T> {

    private String pattern;

    public DateFieldTranslation(String field) {
        super(field);
    }

    public DateFieldTranslation(String field, String pattern) {
        super(field);
        this.pattern = pattern;
    }

    @Override
    public Object translate(T object) {
        Object value = super.translate(object);
        if (StringUtils.isNotBlank(pattern) && value != null && value instanceof Date) {
            return DateUtil.date2String((Date) value, pattern);
        }
        return null;
    }

}
