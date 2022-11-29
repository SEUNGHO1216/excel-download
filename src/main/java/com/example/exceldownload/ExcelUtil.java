package com.example.exceldownload;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Pageable;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Component;
import org.springframework.util.ObjectUtils;
import org.springframework.web.servlet.view.document.AbstractXlsView;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Slf4j
@Component
@RequiredArgsConstructor
public class ExcelUtil {

  private static final int MAX_ROW = 1_040_000;
  private static final int PAGE_SIZE = 10_000;

  private final ProductRepository productRepository;
  private final ProductMapper productMapper;

  public void buildExcelDocument(Map<String, Object> model,
                                 HttpServletRequest request,
                                 HttpServletResponse response) throws Exception {
    long start = System.currentTimeMillis();
    List<String> headerKeys = (List<String>) model.get("headerKeys");
    List<String> widths = (List<String>) model.get("widths");
    String filename = (String) model.get("fileName");


    SXSSFWorkbook sxssfWorkbook = null;
    Pageable pageable = PageRequest.of(0, PAGE_SIZE);
    log.info("page size > {}", pageable.getPageSize());
    Page<ProductDTO> productDTOS = this.getProductList(pageable);
    int totalPages = productDTOS.getTotalPages();
    log.info("totalPages >> {}", totalPages);

    for (int i = 0; i < totalPages; i++) {
      // 헤더에 의해서 이 부분이 어떻게 바뀔지는 보류
      int rowIndex = i * PAGE_SIZE;
      Pageable pageable2 = PageRequest.of(i, PAGE_SIZE);
      Page<ProductDTO> excelDataList = this.getProductList(pageable2);
      log.info("페이징 된 사이즈 >> {}", excelDataList.getContent().size());
      // 헤더 키에 1:1 매핑, 만개의 리스트 == 만개의 로우
      List<Map<String, Object>> headerKeysMap = new ArrayList<>();

      for (ProductDTO excelData : excelDataList) {
        Map<String, Object> tempMap = new HashMap<>();
        tempMap.put("id", excelData.getId());
        tempMap.put("name", excelData.getName());
        tempMap.put("description", excelData.getDescription());
        tempMap.put("price", excelData.getPrice());
        tempMap.put("expireDate", excelData.getExpireDate());

        headerKeysMap.add(tempMap);
      }
      sxssfWorkbook = getWorkBook(headerKeys, widths, headerKeysMap, rowIndex, sxssfWorkbook);
      headerKeysMap.clear(); //초기화
    }

    writeSXSSFWorkbook(sxssfWorkbook, filename, request, response);

    long end = System.currentTimeMillis();
    long gap = end - start;
    log.info("소요시간 >>{} ms", gap);

  }

  public void writeSXSSFWorkbook(SXSSFWorkbook sxssfWorkbook,
                                 String filename,
                                 HttpServletRequest request,
                                 HttpServletResponse response) throws IOException {

    String userAgent = request.getHeader("User-Agent");
    log.info("user agent >> {}", userAgent);
    if (userAgent.contains("Trident") || (userAgent.indexOf("MSIE") > -1)) {
      filename = URLEncoder.encode(filename, "UTF-8").replaceAll("\\+", "%20");
    } else if (userAgent.contains("Chrome")
      || userAgent.contains("Opera")
      || userAgent.contains("Firefox")) {
      filename = new String(filename.getBytes("UTF-8"), "ISO-8859-1");
    }
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

  private SXSSFWorkbook getWorkBook(List<String> headerKeys,
                                    List<String> widths,
                                    List<Map<String, Object>> headerKeysMap,
                                    int rowIndex,
                                    SXSSFWorkbook sxssfWorkbook) throws IOException {
    log.info("get workbook >> {}", sxssfWorkbook);
    SXSSFWorkbook workbook = ObjectUtils.isEmpty(sxssfWorkbook)
      ? new SXSSFWorkbook(-1) : sxssfWorkbook;

    String sheetName = "Sheet" + (rowIndex / MAX_ROW + 1);
    log.info("rowIndx >> {}", rowIndex);
    boolean isNewSheet = ObjectUtils.isEmpty(workbook.getSheet(sheetName));
    log.info("isNewSheet >> {}", isNewSheet);
    Sheet sheet = isNewSheet ? workbook.createSheet(sheetName) : workbook.getSheet(sheetName);

    CellStyle headerStyle = createHeaderStyle(workbook);
    CellStyle bodyStyleLeft = createBodyStyle(workbook, "LEFT");
    CellStyle bodyStyleRight = createBodyStyle(workbook, "RIGHT");
    CellStyle bodyStyleCenter = createBodyStyle(workbook, "CENTER");

    // \r\n을 통해 셀 내 개행
    // 개행을 위해 setWrapText 설정
    bodyStyleLeft.setWrapText(true);
    bodyStyleRight.setWrapText(true);
    bodyStyleCenter.setWrapText(true);

    int columnIndex = 0;

    for (String width : widths) {
      sheet.setColumnWidth(columnIndex++, Integer.parseInt(width) * 256); // 왜 256 곱하지?
    }

    Row row = null;
    Cell cell = null;

    // 매개변수로 받은 rowIdx % MAX_ROW 행부터 이어서 데이터
    int rowNo = rowIndex % MAX_ROW;

    if (isNewSheet) {
      //새로운 시트면 항상 헤더를 새로 만들어줘야 할 것임
      //rowNo은 0부터 시작
      row = sheet.createRow(rowNo);
      columnIndex = 0;

      for (String header : headerKeys) {
        cell = row.createCell(columnIndex++);

        cell.setCellStyle(headerStyle);
        log.info(header);
        cell.setCellValue(header);
      }
    }
    for (Map<String, Object> dataMap : headerKeysMap) {
      columnIndex = 0;
      row = sheet.createRow(++rowNo);

      for (String headerKey : headerKeys) {
        cell = row.createCell(columnIndex++);
        Object value = dataMap.get(headerKey);

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
      }
      // 주기적 flush
      if (rowNo % 100 == 0) {
        ((SXSSFSheet) sheet).flushRows(100);
      }
    }
    return workbook;
  }

  public Page<ProductDTO> getProductList(Pageable pageable) {
    // 엑셀에 저장할 데이터
    return productRepository.findAll(pageable).map(productMapper::toDTO);
  }

  // capitalize the first letter of the field name for retrieving value of the
  // field later
  public static String capitalizeInitialLetter(String s) {
    if (s.length() == 0)
      return s;
    return s.substring(0, 1).toUpperCase() + s.substring(1);
  }

  private CellStyle createHeaderStyle(Workbook workbook) {
    CellStyle headerStyle = createBodyStyle(workbook, "CENTER");
    // 취향에 따라 설정 가능
    headerStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_YELLOW.getIndex());
    headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

    // 가로 세로 정렬 기준
    headerStyle.setAlignment(HorizontalAlignment.CENTER);
    headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

    return headerStyle;
  }

  private CellStyle createBodyStyle(Workbook workbook, String align) {
    CellStyle bodyStyle = workbook.createCellStyle();
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

    return bodyStyle;
  }
}
