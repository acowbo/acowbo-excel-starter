package com.acowbo.excel.controller;


import com.acowbo.excel.annotation.ExcelColumn;
import com.acowbo.excel.model.ExcelTemplate;
import com.acowbo.excel.service.ExcelService;
import com.acowbo.excel.util.ExcelAnnotationUtil;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;
import java.util.Arrays;

/**
 * description: Excel示例控制器
 *
 * @author <a href="https://acowbo.fun">acowbo</a>
 * @version 1.0
 * @since 2025/5/19
 */
@Slf4j
@RestController
@RequestMapping("/excel")
public class ExcelController {

    @Resource
    private ExcelService excelService;

    /**
     * 员工信息实体类
     */
    public static class EmployeeExcel {
        @ExcelColumn(value = "姓名")
        private String name;

        @ExcelColumn(value = "年龄")
        private Integer age;

        @ExcelColumn(value = "手机号", validationType = ExcelColumn.ValidationType.PHONE)
        private String phone;

        @ExcelColumn(value = "邮箱")
        private String email;

        @ExcelColumn(value = "身份证号", validationType = ExcelColumn.ValidationType.ID_CARD)
        private String idCard;

        @ExcelColumn(value = "银行卡号", validationType = ExcelColumn.ValidationType.BANK_CARD)
        private String bankCard;

        @ExcelColumn(value = "状态")
        private String status;
    }

    @GetMapping("/template")
    public void downloadTemplate(HttpServletResponse response) {
        // 使用注解工具类创建模板
        ExcelTemplate template = ExcelAnnotationUtil.createTemplateFromClass(EmployeeExcel.class);
        
        // 添加下拉选项
        template.addDropdownOptions("status", Arrays.asList("在职", "离职", "休假"));
        
        // 下载模板
        excelService.downloadTemplate(response, template, "员工信息模板.xlsx");
    }

    @PostMapping("/import")
    public String importData(@RequestPart("file") MultipartFile file) {
        // 使用注解工具类创建模板
        ExcelTemplate template = ExcelAnnotationUtil.createTemplateFromClass(EmployeeExcel.class);
        
        // 添加下拉选项
        template.addDropdownOptions("status", Arrays.asList("在职", "离职", "休假"));

        // 导入数据
        return excelService.importData(file, template, dataList -> {
            // 这里处理导入的数据
            log.info("导入数据：{}", dataList);
            return "导入成功，共导入" + dataList.size() + "条数据";
        });
    }
} 