package com.example.exceldownload.util;

import lombok.Getter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;

@Getter
public enum ExcelStyle {
  HEADER_STYLE(),
  BODY_STYLE;

  private CellStyle cellStyle;

  ExcelStyle(CellStyle cellStyle){
    this.cellStyle = cellStyle;
  }

  ExcelStyle() {

  }

  public CellStyle getCellStyle(Workbook workbook, ExcelStyle excelStyle){
    String align = "CENTER";
    CellStyle bodyStyle = workbook.createCellStyle();
    if(excelStyle.equals(ExcelStyle.BODY_STYLE)) {

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
    }
    return bodyStyle;
  }
}
