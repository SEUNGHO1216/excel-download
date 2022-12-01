package com.example.exceldownload.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Slf4j
public class ExcelHandler {

  private SXSSFWorkbook sxssfWorkbook;
  private ExcelUtil excelUtil;

  public ExcelHandler() {
    sxssfWorkbook = new SXSSFWorkbook();
    excelUtil = new ExcelUtil();
  }

  public void excelDownload(List<String> headerKeys,
                            List<String> widths,
                            List<Map<String, Object>> headerKeysMap,
                            int rowIndex,
                            boolean isLast,
                            HttpServletRequest request,
                            HttpServletResponse response) throws IOException {
    log.info("excel util 주소값 >> {}", excelUtil);
    sxssfWorkbook = excelUtil.getWorkBook(headerKeys, widths, headerKeysMap, rowIndex, sxssfWorkbook);
    if(isLast){
      String filename = "Excel_Download_Test";
      excelUtil.writeSXSSFWorkbook(sxssfWorkbook, filename, request, response);
      sxssfWorkbook = null;
      log.info("=====엑셀 다운로드 완료=====");
    }
  }

  public List<String> getKeys(Map<String, Object> dataMap, List<String> keys){
    dataMap.forEach((key,value)->{
      if(value instanceof LinkedHashMap){
        Map<String, Object> map = (LinkedHashMap) value;
        getKeys(map, keys);
      }
      keys.add(key);
    });
    return keys;
  }
}
