package com.woter.fact.excel.core;

import com.woter.fact.excel.exception.ExcelException;
import org.apache.commons.lang3.reflect.FieldUtils;

import java.util.Map;

/**
 * @author woter
 * @date 2017-9-29 下午6:52:05
 */
public class DefaultFieldTranslation<T> implements FieldTranslation<T> {

    private String field;

    public DefaultFieldTranslation(String field) {
        this.field = field;
    }


    public Object translate(T object) {
        try {
            if (object instanceof Map) {
                Map map = (Map) object;
                return map.get(field);
            } else {
                return FieldUtils.readDeclaredField(object, field, true);
            }
        } catch (IllegalAccessException e) {
            throw new ExcelException(e);
        }
    }

}
