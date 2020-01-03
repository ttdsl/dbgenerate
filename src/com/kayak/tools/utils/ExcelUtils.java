package com.kayak.tools.utils;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.HashMap;
import java.util.concurrent.locks.ReentrantLock;

/**
 * Excel工具类
 * @author jintao
 *
 */
public class ExcelUtils {
    private static ReentrantLock reentrantLock = new ReentrantLock();
    private static HashMap<String,CellStyle> cellStyles = new HashMap<String,CellStyle>();//存放所有会使用到的Excel样式

    public static CellStyle getCellStyle(String cellStyle){
        return cellStyles.get(cellStyle);
    }
    public static void cteateCell(HSSFWorkbook wb, HSSFRow row, int col, String val, CellStyle style) {
        final ReentrantLock lock = reentrantLock;
        synchronized (lock){
            HSSFCell cell = row.createCell(col);
            cell.setCellStyle(style);
            cell.setCellValue(val);
            // 使用HSSFCell.CELL_TYPE_STRING 在多线程调用出现了问题,报空指针异常
            cell.setCellType(1);
        }
    }

    public static void doCreateCellStyle(HSSFWorkbook hssfWorkbook){
        CellStyle log = ExcelUtils.createCellStyle(hssfWorkbook, "LOG");//LOG页签头部样式
        CellStyle head_name = ExcelUtils.createCellStyle(hssfWorkbook, "HEAD_NAME");//index页签表名样式
        CellStyle th_left = ExcelUtils.createCellStyle(hssfWorkbook, "TH_LEFT");//首行左侧
        CellStyle th_right = ExcelUtils.createCellStyle(hssfWorkbook, "TH_RIGHT");//首行右侧
        CellStyle td_center_left = ExcelUtils.createCellStyle(hssfWorkbook, "TD_CENTER_LEFT");//中间左侧
        CellStyle td_center_right = ExcelUtils.createCellStyle(hssfWorkbook, "TD_CENTER_RIGHT");//中间右侧
        CellStyle td_bottom_left = ExcelUtils.createCellStyle(hssfWorkbook, "TD_BOTTOM_LEFT");//底层左侧
        CellStyle td_bottom_right = ExcelUtils.createCellStyle(hssfWorkbook, "TD_BOTTOM_RIGHT");//底层右侧
        CellStyle def = ExcelUtils.createCellStyle(hssfWorkbook, "DEFAULT");//默认样式
        CellStyle index = ExcelUtils.createCellStyle(hssfWorkbook,"INDEX");//索引头样式
        cellStyles.put("log",log);
        cellStyles.put("head_name",head_name);
        cellStyles.put("th_left",th_left);
        cellStyles.put("th_right",th_right);
        cellStyles.put("td_center_left",td_center_left);
        cellStyles.put("td_center_right",td_center_right);
        cellStyles.put("td_bottom_left",td_bottom_left);
        cellStyles.put("td_bottom_right",td_bottom_right);
        cellStyles.put("def",def);
        cellStyles.put("index",index);
    }

    /**
     * @param partten Excel表格位置 exp HEAD TH TD TD_BOTTOM
     * @return
     */
    public static CellStyle createCellStyle(HSSFWorkbook hssfWorkbook, String partten){
        CellStyle style = hssfWorkbook.createCellStyle();
        HSSFFont font = hssfWorkbook.createFont();
        font.setFontHeight((short)10);
        font.setFontName("宋体");
        switch(partten){
            case "LOG":
                font.setFontHeightInPoints((short) 14);
                font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//粗体显示
                style.setAlignment(HSSFCellStyle.ALIGN_CENTER_SELECTION);
                HSSFPalette palette_log = hssfWorkbook.getCustomPalette();
                palette_log.setColorAtIndex(HSSFColor.LIME.index, (byte) 255, (byte) 255, (byte) 71);
                style.setFillForegroundColor(HSSFColor.LIME.index);
                style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                style.setFont(font);
                return style;
            case "HEAD_NAME":
                font.setFontName("楷体");
                font.setFontHeightInPoints((short) 12);
                style.setAlignment(HSSFCellStyle.ALIGN_LEFT);
                style.setFont(font);
                return style;
            case "HEAD_COMMENT":
                font.setFontHeightInPoints((short) 12);
                style.setAlignment(HSSFCellStyle.ALIGN_CENTER_SELECTION);
                style.setFont(font);
                return style;
            case "INDEX":
                font.setFontName("楷体");
                font.setFontHeightInPoints((short) 12);
                style.setAlignment(HSSFCellStyle.ALIGN_LEFT);
                font.setUnderline((byte) 1);
                font.setColor(HSSFColor.BLUE.index);
                style.setFont(font);
                return style;
            case "TH_LEFT":
                font.setFontHeightInPoints((short) 10);
                //单元格上方 中等 蓝色边框 下方、左侧、右侧 虚线蓝色边框
                style.setBorderTop(CellStyle.BORDER_MEDIUM);
                style.setTopBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderRight(CellStyle.BORDER_DASHED);
                style.setRightBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderBottom(CellStyle.BORDER_DASHED);
                style.setBottomBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderLeft(CellStyle.BORDER_DASHED);
                style.setLeftBorderColor(IndexedColors.BLUE.getIndex());

                style.setFillForegroundColor(HSSFColor.LIME.index);
                style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                HSSFPalette palette = hssfWorkbook.getCustomPalette();
                palette.setColorAtIndex(HSSFColor.LIME.index, (byte) 153, (byte) 204, (byte) 255);
                style.setFont(font);
                return style;
            case "TH_RIGHT":
                font.setFontHeightInPoints((short) 10);
                //单元格上方和右方 中等 蓝色边框 下方虚线蓝色边框 左侧 虚线蓝色边框
                style.setBorderTop(CellStyle.BORDER_MEDIUM);
                style.setTopBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderRight(CellStyle.BORDER_MEDIUM);
                style.setRightBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderBottom(CellStyle.BORDER_DASHED);
                style.setBottomBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderLeft(CellStyle.BORDER_DASHED);
                style.setLeftBorderColor(IndexedColors.BLUE.getIndex());

                //设置背景颜色
                style.setFillForegroundColor(HSSFColor.LIME.index);
                //solid 填充  foreground  前景色
                style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                HSSFPalette palette_right = hssfWorkbook.getCustomPalette();
                palette_right.setColorAtIndex(HSSFColor.LIME.index, (byte) 153, (byte) 204, (byte) 255);
                style.setFont(font);
                return style;
            case "TD_CENTER_LEFT":
                font.setFontHeightInPoints((short) 10);
                //单元格全部为蓝色虚线
                style.setBorderTop(CellStyle.BORDER_DASHED);
                style.setTopBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderBottom(CellStyle.BORDER_DASHED);
                style.setBottomBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderLeft(CellStyle.BORDER_DASHED);
                style.setLeftBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderRight(CellStyle.BORDER_DASHED);
                style.setRightBorderColor(IndexedColors.BLUE.getIndex());
                style.setAlignment(HSSFCellStyle.ALIGN_LEFT);
                style.setFont(font);
                return style;
            case "TD_CENTER_RIGHT":
                font.setFontHeightInPoints((short) 10);
                //单元格右侧为中等实线 其余为虚线
                style.setBorderTop(CellStyle.BORDER_DASHED);
                style.setTopBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderBottom(CellStyle.BORDER_DASHED);
                style.setBottomBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderLeft(CellStyle.BORDER_DASHED);
                style.setLeftBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderRight(CellStyle.BORDER_MEDIUM);
                style.setRightBorderColor(IndexedColors.BLUE.getIndex());
                style.setAlignment(HSSFCellStyle.ALIGN_LEFT);
                style.setFont(font);
                return style;
            case "TD_BOTTOM_LEFT":
                font.setFontHeightInPoints((short) 10);
                //单元格上、左、右全部为 蓝色虚线  下侧 蓝色中等实线
                style.setBorderTop(CellStyle.BORDER_DASHED);
                style.setTopBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderLeft(CellStyle.BORDER_DASHED);
                style.setLeftBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderRight(CellStyle.BORDER_DASHED);
                style.setRightBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderBottom(CellStyle.BORDER_MEDIUM);
                style.setBottomBorderColor(IndexedColors.BLUE.getIndex());
                style.setAlignment(HSSFCellStyle.ALIGN_LEFT);
                style.setFont(font);
                return style;
            case "TD_BOTTOM_RIGHT":
                font.setFontHeightInPoints((short) 10);
                //单元格上、左为蓝色虚线  下侧、右侧 蓝色中等实线
                style.setBorderTop(CellStyle.BORDER_DASHED);
                style.setTopBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderLeft(CellStyle.BORDER_DASHED);
                style.setLeftBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderRight(CellStyle.BORDER_MEDIUM);
                style.setRightBorderColor(IndexedColors.BLUE.getIndex());

                style.setBorderBottom(CellStyle.BORDER_MEDIUM);
                style.setBottomBorderColor(IndexedColors.BLUE.getIndex());
                style.setAlignment(HSSFCellStyle.ALIGN_LEFT);
                style.setFont(font);
                return style;
        }
        font.setFontHeightInPoints((short) 10);
        style.setFont(font);
        style.setAlignment(HSSFCellStyle.ALIGN_LEFT);
        return style;
    }
}
