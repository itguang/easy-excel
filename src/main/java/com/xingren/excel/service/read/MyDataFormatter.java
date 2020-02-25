package com.xingren.excel.service.read;

import org.apache.poi.ss.formula.ConditionalFormattingEvaluator;
import org.apache.poi.ss.usermodel.*;

import java.util.Date;

import static com.xingren.excel.ExcelConstant.DEFAULT_DATE_PATTREN;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;

/**
 * @author guang
 * @since 2020/2/25 4:08 下午
 */
public class MyDataFormatter extends DataFormatter {

    @Override
    public String formatCellValue(Cell cell, FormulaEvaluator evaluator, ConditionalFormattingEvaluator cfEvaluator) {
        if (cell == null) {
            return "";
        }

        CellType cellType = cell.getCellTypeEnum();
        if (cellType == CellType.FORMULA) {
            if (evaluator == null) {
                return cell.getCellFormula();
            }
            cellType = evaluator.evaluateFormulaCellEnum(cell);
        }

        // 日期处理,转字符串 yyyy/MM/dd
        if (cellType.equals(NUMERIC) && DateUtil.isCellDateFormatted(cell, cfEvaluator)) {
            Date cellDateCellValue = cell.getDateCellValue();
            // todo#guang 由于不知道怎么获取 日期格式 Cell 的原始字符串,所以目前只能转成默认格式的日期
            String formatDate = com.xingren.excel.util.DateUtil.formatDate(cellDateCellValue, DEFAULT_DATE_PATTREN);
            return formatDate;
        }
        return super.formatCellValue(cell, evaluator, cfEvaluator);
    }
}
