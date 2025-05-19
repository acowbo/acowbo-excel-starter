package com.acowbo.excel.builder;


import com.acowbo.excel.annotation.ExcelColumn;
import com.acowbo.excel.model.ExcelTemplate;
import com.acowbo.excel.model.ExcelValidation;


import java.util.*;

/**
 * description: Excel模板构建器
 *
 * @author <a href="https://acowbo.fun">acowbo</a>
 * @version 1.0
 * @since 2025/5/19
 */
public class ExcelTemplateBuilder {
    private final Map<String, Object> templateData = new HashMap<>();
    private final LinkedHashMap<String, String> headers = new LinkedHashMap<>();
    private final Map<String, Integer> columnIndices = new HashMap<>();
    private final List<ExcelValidation> validations = new ArrayList<>();
    private final Map<String, List<Map<String, Object>>> dropdownOptions = new HashMap<>();

    /**
     * 添加列
     */
    public ExcelTemplateBuilder addColumn(String fieldName, String displayName) {
        headers.put(fieldName, displayName);
        columnIndices.put(fieldName, headers.size() - 1);
        return this;
    }

    /**
     * 添加列（使用注解）
     */
    public ExcelTemplateBuilder addColumn(ExcelColumn column) {
        String fieldName = column.value();
        headers.put(fieldName, fieldName);
        columnIndices.put(fieldName, headers.size() - 1);

        // 添加验证
        if (column.validationType() != ExcelColumn.ValidationType.NONE) {
            ExcelValidation validation = ExcelValidation.builder()
                    .type(column.validationType().name().toLowerCase())
                    .title(fieldName)
                    .fieldName(fieldName)
                    .minLength(column.minLength())
                    .maxLength(column.maxLength())
                    .build();
            validations.add(validation);
        }

        // 添加下拉选项
        if (!column.dropdownKey().isEmpty()) {
            dropdownOptions.put(fieldName, new ArrayList<>());
        }

        return this;
    }

    /**
     * 添加下拉选项
     */
    public ExcelTemplateBuilder addDropdownOption(String fieldName, String optionName) {
        List<Map<String, Object>> options = dropdownOptions.computeIfAbsent(fieldName, k -> new ArrayList<>());
        Map<String, Object> option = new HashMap<>();
        option.put("name", optionName);
        options.add(option);
        return this;
    }

    /**
     * 添加下拉选项列表
     */
    public ExcelTemplateBuilder addDropdownOptions(String fieldName, List<String> optionNames) {
        List<Map<String, Object>> options = dropdownOptions.computeIfAbsent(fieldName, k -> new ArrayList<>());
        optionNames.forEach(name -> {
            Map<String, Object> option = new HashMap<>();
            option.put("name", name);
            options.add(option);
        });
        return this;
    }

    /**
     * 添加验证
     */
    public ExcelTemplateBuilder addValidation(ExcelValidation validation) {
        validations.add(validation);
        return this;
    }

    /**
     * 构建模板
     */
    public ExcelTemplate build() {
        templateData.put("headers", headers);
        templateData.put("columnIndices", columnIndices);

        // 添加验证
        validations.forEach(validation -> {
            Map<String, Object> validationMap = new HashMap<>();
            validationMap.put("type", validation.getType());
            validationMap.put("title", validation.getTitle());
            validationMap.put("fieldName", validation.getFieldName());
            validationMap.put("minLength", validation.getMinLength());
            validationMap.put("maxLength", validation.getMaxLength());
            templateData.put(validation.getFieldName() + "Validation", validationMap);
        });

        // 添加下拉选项
        dropdownOptions.forEach((fieldName, options) -> {
            if (!options.isEmpty()) {
                templateData.put(fieldName + "List", options);
            }
        });

        return new ExcelTemplate(templateData);
    }

    /**
     * 创建构建器
     */
    public static ExcelTemplateBuilder create() {
        return new ExcelTemplateBuilder();
    }
} 