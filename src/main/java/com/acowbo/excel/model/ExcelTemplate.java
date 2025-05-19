package com.acowbo.excel.model;

import lombok.Data;
import java.util.*;

/**
 * description: Excel模板模型
 *
 * @author <a href="https://acowbo.fun">acowbo</a>
 * @version 1.0
 * @since 2025/5/19
 */
@Data
public class ExcelTemplate {
    private final Map<String, Object> templateData;

    public ExcelTemplate(Map<String, Object> templateData) {
        this.templateData = templateData;
    }

    /**
     * 获取表头映射
     */
    @SuppressWarnings("unchecked")
    public Map<String, String> getHeaders() {
        return (Map<String, String>) templateData.get("headers");
    }

    /**
     * 获取列索引映射
     */
    @SuppressWarnings("unchecked")
    public Map<String, Integer> getColumnIndices() {
        return (Map<String, Integer>) templateData.get("columnIndices");
    }

    /**
     * 获取验证规则
     */
    @SuppressWarnings("unchecked")
    public List<Map<String, Object>> getValidations() {
        List<Map<String, Object>> validations = new ArrayList<>();
        templateData.forEach((key, value) -> {
            if (key.endsWith("Validation") && value instanceof Map) {
                validations.add((Map<String, Object>) value);
            }
        });
        return validations;
    }

    /**
     * 获取下拉选项
     */
    @SuppressWarnings("unchecked")
    public Map<String, List<Map<String, Object>>> getDropdownOptions() {
        Map<String, List<Map<String, Object>>> options = new HashMap<>();
        templateData.forEach((key, value) -> {
            if (key.endsWith("List") && value instanceof List) {
                String fieldName = key.substring(0, key.length() - 4);
                options.put(fieldName, (List<Map<String, Object>>) value);
            }
        });
        return options;
    }

    /**
     * 添加下拉选项
     */
    public void addDropdownOptions(String fieldName, List<String> options) {
        List<Map<String, Object>> optionList = new ArrayList<>();
        options.forEach(option -> {
            Map<String, Object> optionMap = new HashMap<>();
            optionMap.put("name", option);
            optionList.add(optionMap);
        });
        templateData.put(fieldName + "List", optionList);
    }

    /**
     * 添加验证规则
     */
    public void addValidation(String type, String title, String fieldName) {
        Map<String, Object> validation = new HashMap<>();
        validation.put("type", type);
        validation.put("title", title);
        validation.put("fieldName", fieldName);
        templateData.put(fieldName + "Validation", validation);
    }

    /**
     * 移除验证规则
     */
    public void removeValidation(String fieldName) {
        templateData.remove(fieldName + "Validation");
    }

    /**
     * 移除下拉选项
     */
    public void removeDropdownOptions(String fieldName) {
        templateData.remove(fieldName + "List");
    }
} 