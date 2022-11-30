package com.example.exceldownload.util;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.util.ObjectUtils;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

@Slf4j
@Component
@RequiredArgsConstructor
public class ExcelUtil {

  private static final int MAX_ROW = 1_040_000;

  public SXSSFWorkbook getWorkBook(List<String> headerKeys,
                                   List<String> widths,
                                   List<Map<String, Object>> headerKeysMap,
                                   int rowIndex,
                                   SXSSFWorkbook sxssfWorkbook) throws IOException {
    log.info("get workbook >> {}", sxssfWorkbook);
    SXSSFWorkbook workbook = ObjectUtils.isEmpty(sxssfWorkbook)
      ? new SXSSFWorkbook(-1) : sxssfWorkbook;

    String sheetName = "TestSheet" + (rowIndex / MAX_ROW + 1);
    log.info("rowIndx >> {}", rowIndex);
    boolean isNewSheet = ObjectUtils.isEmpty(workbook.getSheet(sheetName));
    log.info("isNewSheet >> {}", isNewSheet);
    Sheet sheet = isNewSheet ? workbook.createSheet(sheetName) : workbook.getSheet(sheetName);

    int columnIndex = 0;

    for (String width : widths) {
      sheet.setColumnWidth(columnIndex++, Integer.parseInt(width) * 256); // 왜 256 곱하지?
    }

    Row row = null;
    Cell cell = null;
    CellStyle cellStyle = workbook.createCellStyle();

    // 매개변수로 받은 rowIndex % MAX_ROW 행부터 이어서 데이터 입력
    int rowNo = rowIndex % MAX_ROW;
    if (isNewSheet) {
      // 새로운 시트면 항상 헤더를 새로 만들어줘야 할 것임 , rowNo은 0부터 시작
      row = sheet.createRow(rowNo);
      // 헤더 == 칼럼
      columnIndex = 0;
      // header cell 입력
      for (String header : headerKeys) {
        log.info("헤더 >> {}", header);
        cell = row.createCell(columnIndex++);
        cell.setCellValue(header);
        cell.setCellStyle(ExcelStyle.HEADER_STYLE.getHeaderCellStyle(cellStyle));
      }
      cellStyle = workbook.createCellStyle();
    }
    // body cell 입력
    for (Map<String, Object> excelValueMap : headerKeysMap) {
      columnIndex = 0;
      row = sheet.createRow(++rowNo);

      for (String headerKey : headerKeys) {
        cell = row.createCell(columnIndex++);
        Object value = excelValueMap.get(headerKey);
        _cellWrite(cell, value, cellStyle);
      }
      // 주기적 flush
      if (rowNo % 100 == 0) {
        ((SXSSFSheet) sheet).flushRows(100);
      }
    }
    return workbook;
  }
  public void writeSXSSFWorkbook(SXSSFWorkbook sxssfWorkbook,
                                 String filename,
                                 HttpServletRequest request,
                                 HttpServletResponse response) throws IOException {

    // 브라우저별 인코딩
    String userAgent = request.getHeader("User-Agent");
    log.info("user agent >> {}", userAgent);
    if (userAgent.contains("Trident") || (userAgent.indexOf("MSIE") > -1)) {
      filename = URLEncoder.encode(filename, "UTF-8").replaceAll("\\+", "%20");
    } else if (userAgent.contains("Chrome")
      || userAgent.contains("Opera")
      || userAgent.contains("Firefox")) {
      filename = new String(filename.getBytes("UTF-8"), "ISO-8859-1");
    }
    // output stream 을 열어서 엑셀 파일로 변환
    try {
      filename = filename + System.currentTimeMillis();
      log.info("filename >> {}", filename);
      response.setContentType("application/vnd.ms-excel");
      response.setHeader("Content-Disposition", "attachment;filename=" + filename + ".xlsx");
      ServletOutputStream outputStream = response.getOutputStream();
      outputStream.flush();
      sxssfWorkbook.write(outputStream);
      outputStream.flush();
      outputStream.close();
    } catch (Exception e) {
      log.error("[SxssfExcelView] error message: {}", e.getMessage());
    } finally {
      if (!ObjectUtils.isEmpty(sxssfWorkbook)) {
        log.info("sxssfWorkBook check (true/false) >> {}", ObjectUtils.isEmpty(sxssfWorkbook));
        sxssfWorkbook.dispose();
        sxssfWorkbook.close();
      }
    }
  }

  private Cell _cellWrite(Cell cell, Object value, CellStyle cellStyle){
    cell.setCellStyle(ExcelStyle.BODY_STYLE.getBodyCellStyle(cellStyle));
    // 숫자 정밀도 보장
    if (value instanceof BigDecimal) {
      cell.setCellValue(((BigDecimal) value).toString());
    } else if (value instanceof Double) {
      cell.setCellValue(((Double) value).toString());
    } else if (value instanceof Long) {
      cell.setCellValue(((Long) value).toString());
    } else if (value instanceof Integer) {
      cell.setCellValue(((Integer) value).toString());
    } else if (value instanceof LocalDateTime) {
      String date = ((LocalDateTime) value).format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
      cell.setCellValue(date);
    } else {
      cell.setCellValue((String) value);
    }

    return cell;
  }

}
