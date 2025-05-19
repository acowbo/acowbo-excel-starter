package com.acowbo.excel.util;

import com.acowbo.excel.config.ExcelTemplateConfig;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.Map;

/**
 * description: Excel验证工具类
 *
 * @author <a href="https://acowbo.fun">acowbo</a>
 * @version 1.0
 * @since 2025/5/19
 */
public class ExcelValidationUtil {

    /**
     * 添加数据验证
     *
     * @param helper 数据验证帮助类
     * @param sheet 工作表
     * @param colIndex 列索引
     * @param validationConfig 验证配置
     * @param config Excel模板配置
     */
    public static void addValidation(DataValidationHelper helper, Sheet sheet, int colIndex,
                                     Map<String, Object> validationConfig, ExcelTemplateConfig config) {
        String type = validationConfig.get("type").toString();
        String title = validationConfig.get("title").toString();
        int minLength = (int) validationConfig.getOrDefault("minLength", -1);
        int maxLength = (int) validationConfig.getOrDefault("maxLength", -1);

        DataValidationConstraint constraint;
        switch (type) {
            case "phone":
                constraint = helper.createTextLengthConstraint(DataValidationConstraint.OperatorType.EQUAL, "11", null);
                addValidationWithStyle(helper, sheet, colIndex, constraint, title, "请输入11位手机号", config);
                break;
            case "idCard":
                constraint = helper.createTextLengthConstraint(DataValidationConstraint.OperatorType.EQUAL, "18", null);
                addValidationWithStyle(helper, sheet, colIndex, constraint, title, "请输入18位身份证号", config);
                break;
            case "bankCard":
                constraint = helper.createTextLengthConstraint(DataValidationConstraint.OperatorType.BETWEEN, "16", "19");
                addValidationWithStyle(helper, sheet, colIndex, constraint, title, "请输入16-19位银行卡号", config);
                break;
            case "email":
                constraint = helper.createCustomConstraint("ISEMAIL()");
                addValidationWithStyle(helper, sheet, colIndex, constraint, title, "请输入正确的邮箱地址", config);
                break;
            case "number":
                constraint = helper.createNumericConstraint(
                        DataValidationConstraint.ValidationType.INTEGER,
                        DataValidationConstraint.OperatorType.GREATER_OR_EQUAL,
                        "0",
                        null
                );
                addValidationWithStyle(helper, sheet, colIndex, constraint, title, "请输入数字", config);
                break;
            case "length":
                if (minLength > 0 && maxLength > 0) {
                    constraint = helper.createTextLengthConstraint(
                            DataValidationConstraint.OperatorType.BETWEEN,
                            String.valueOf(minLength),
                            String.valueOf(maxLength)
                    );
                    addValidationWithStyle(helper, sheet, colIndex, constraint, title,
                            String.format("请输入%d-%d位字符", minLength, maxLength), config);
                }
                break;
        }
    }

    /**
     * 添加带样式的数据验证
     */
    private static void addValidationWithStyle(DataValidationHelper helper, Sheet sheet, int colIndex,
                                               DataValidationConstraint constraint, String title, String message,
                                               ExcelTemplateConfig config) {
        CellRangeAddressList rangeList = new CellRangeAddressList(1, config.getMaxDropdownRows(), colIndex, colIndex);
        DataValidation validation = helper.createValidation(constraint, rangeList);
        validation.setShowErrorBox(config.isEnableErrorPrompt());
        validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        validation.createErrorBox("格式错误", message);
        validation.setShowPromptBox(config.isEnableWarningPrompt());
        validation.createPromptBox(title, message);
        sheet.addValidationData(validation);
    }
} 