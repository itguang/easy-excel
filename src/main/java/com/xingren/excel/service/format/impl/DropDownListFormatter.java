package com.xingren.excel.service.format.impl;

import com.xingren.excel.pojo.TemplateRowFormat;
import com.xingren.excel.service.format.IFormatter;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

/**
 * @author ellin
 * @since 2020/08/03
 */
public class DropDownListFormatter implements IFormatter {

    @Override
    public void format(Sheet sheet, TemplateRowFormat format) {
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = helper.createExplicitListConstraint(format.getData());
        CellRangeAddressList regions = new CellRangeAddressList();
        regions.addCellRangeAddress(format.getRange());
        DataValidation validation = helper.createValidation(constraint, regions);

        if (validation instanceof XSSFDataValidation) {
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
        } else {
            validation.setSuppressDropDownArrow(false);
        }

        sheet.addValidationData(validation);
    }

}
