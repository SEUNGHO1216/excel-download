package com.example.exceldownload.util;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Component;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Slf4j
@Component
@RequiredArgsConstructor
public class ExcelHandler {

  private static SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook();
  private final ExcelUtil excelUtil;

  public SXSSFWorkbook makeSXSSFWorkbook(){
    return sxssfWorkbook;
  }

  public synchronized void excelDownload(List<String> headerKeys,
                            List<String> widths,
                            List<Map<String, Object>> headerKeysMap,
                            int rowIndx,
                            boolean flag,
                            HttpServletRequest request,
                            HttpServletResponse response) throws IOException {

    sxssfWorkbook = excelUtil.getWorkBook(headerKeys, widths, headerKeysMap, rowIndx, sxssfWorkbook);
    if(flag){
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
