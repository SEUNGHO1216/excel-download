package com.example.exceldownload.util;

import lombok.NoArgsConstructor;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

@NoArgsConstructor
public enum ExcelStyle {
  HEADER_STYLE(),
  BODY_STYLE;

  private CellStyle cellStyle;

  ExcelStyle(CellStyle cellStyle) {
    this.cellStyle = cellStyle;
  }

  public CellStyle getHeaderCellStyle(CellStyle cellStyle) {
    String align = "CENTER";
    CellStyle headerStyle = getBodyCellStyle(cellStyle);

    // 취향에 따라 설정 가능
    headerStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_YELLOW.getIndex());
    headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

    // 가로 세로 정렬 기준
    headerStyle.setAlignment(HorizontalAlignment.CENTER);
    headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

    return headerStyle;
  }

  public CellStyle getBodyCellStyle(CellStyle bodyStyle) {
    String align = "CENTER";

    // 취향에 따라 설정 가능
    bodyStyle.setBorderTop(BorderStyle.THIN);
    bodyStyle.setBorderBottom(BorderStyle.THIN);
    bodyStyle.setBorderLeft(BorderStyle.THIN);
    bodyStyle.setBorderRight(BorderStyle.THIN);
    bodyStyle.setVerticalAlignment(VerticalAlignment.CENTER);

    if (!StringUtils.isEmpty(align)) {
      if ("LEFT".equals(align)) {
        bodyStyle.setAlignment(HorizontalAlignment.LEFT);
      } else if ("RIGHT".equals(align)) {
        bodyStyle.setAlignment(HorizontalAlignment.RIGHT);
      } else {
        bodyStyle.setAlignment(HorizontalAlignment.CENTER);
      }
    }
    // \r\n을 통해 셀 내 개행, 개행을 위해 setWrapText 설정
    bodyStyle.setWrapText(true);
    return bodyStyle;
  }
}
