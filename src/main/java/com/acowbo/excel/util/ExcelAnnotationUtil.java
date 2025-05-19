package com.acowbo.excel.util;



import com.acowbo.excel.annotation.ExcelColumn;
import com.acowbo.excel.builder.ExcelTemplateBuilder;
import com.acowbo.excel.model.ExcelTemplate;
import com.acowbo.excel.model.ExcelValidation;

import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.List;

/**
 * description: Excel注解工具类
 *
 * @author <a href="https://acowbo.fun">acowbo</a>
 * @version 1.0
 * @since 2025/5/19
 */
public class ExcelAnnotationUtil {
    
    /**
     * 从实体类创建Excel模板
     *
     * @param clazz 实体类Class
     * @return Excel模板
     */
    public static ExcelTemplate createTemplateFromClass(Class<?> clazz) {
        ExcelTemplateBuilder builder = ExcelTemplateBuilder.create();
        
        // 获取所有字段
        Field[] fields = clazz.getDeclaredFields();
        
        // 遍历字段，读取注解信息
        for (Field field : fields) {
            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
            if (annotation != null) {
                // 添加列
                builder.addColumn(field.getName(), annotation.value());
                
                // 添加验证
                if (annotation.validationType() != ExcelColumn.ValidationType.NONE) {
                    builder.addValidation(ExcelValidation.builder()
                            .type(annotation.validationType().getValue())
                            .title(annotation.value())
                            .fieldName(field.getName())
                            .minLength(annotation.minLength())
                            .maxLength(annotation.maxLength())
                            .build());
                }
            }
        }
        
        return builder.build();
    }

    /**
     * 获取实体类的字段名
     */
    public static String getFieldName(Class<?> clazz, String fieldName) {
        try {
            return clazz.getDeclaredField(fieldName).getName();
        } catch (NoSuchFieldException e) {
            throw new RuntimeException("字段不存在: " + fieldName, e);
        }
    }
} 