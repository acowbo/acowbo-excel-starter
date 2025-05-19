package com.acowbo.excel.model;

import lombok.Builder;
import lombok.Data;

/**
 * description: Excel验证模型
 *
 * @author <a href="https://acowbo.fun">acowbo</a>
 * @version 1.0
 * @since 2025/5/19
 */
@Data
@Builder
public class ExcelValidation {
    /**
     * 验证类型
     */
    private String type;

    /**
     * 验证标题
     */
    private String title;

    /**
     * 字段名
     */
    private String fieldName;

    /**
     * 最小长度
     */
    private int minLength;

    /**
     * 最大长度
     */
    private int maxLength;
} 