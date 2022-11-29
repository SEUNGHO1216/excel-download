package com.example.exceldownload;

import com.fasterxml.jackson.annotation.JsonTypeInfo;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Pageable;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;

@Slf4j
@Service
@RequiredArgsConstructor
public class ProductService {

  private final ProductRepository productRepository;
  private final ProductMapper productMapper;
//  private final ExcelUtil excelUtil;
  private final ExcelHandler excelHandler;
  private final ObjectMapper objectMapper;


//  public void getExcelList(HttpServletRequest request, HttpServletResponse response) throws Exception {
//    long start = System.currentTimeMillis();
//    int page = 0;
//    int size = 10000;
//    Page<ProductDTO> productDTOS = getProductList(PageRequest.of(page, size));
//    @SuppressWarnings("unchecked")
//    Class<ProductDTO> productDTOClass = (Class<ProductDTO>) productDTOS.getContent().get(0).getClass();
//    Field[] productFields = productDTOClass.getDeclaredFields();
//    int columnLength = productFields.length;
//
//    List<String> headerKeys = new ArrayList<>();
//    for (int i = 0; i < columnLength; i++) {
//      headerKeys.add(productFields[i].getName());
//    }
//
////    List<String> headerKeys = List.of("id", "name", "description", "price", "expireDate"); // 이 자체가 헤더 키이자 헤더이다
//    // 헤더당 셀 너비 설정
//    List<String> widths = Arrays.asList("10", "20", "50", "15", "20");
//    // 파일명 설정
//    String fileName = "EXAMPLE_EXCEL";
//    SXSSFWorkbook sxssfWorkbook = null;
//    int totalPages = productDTOS.getTotalPages();
//    for (int i = 0; i < totalPages; i++) {
//      Page<ProductDTO> pagedExcelData = getProductList(PageRequest.of(i, size));
//      int rowIndex = i * size;
//      List<Map<String, Object>> headerKeysMap = new ArrayList<>();
//
//      // 헤더 키에 1:1 매핑, 만개의 리스트 == 만개의 로우
//      for (ProductDTO excelData : pagedExcelData) {
//        Map<String, Object> tempMap = new HashMap<>();
//        for (Field field : productFields) {
//          String fieldName = field.getName();
//          Method method = null; // reflect
//          try {
//            method = productDTOClass.getMethod("get" + ExcelUtil.capitalizeInitialLetter(fieldName));
//          } catch (NoSuchMethodException e) {
//            method = productDTOClass.getMethod("get" + fieldName);
//          }
//          Object value = method.invoke(excelData, (Object[]) null);
//          tempMap.put(fieldName, value);
//        }
//        headerKeysMap.add(tempMap);
//      }
//
//      Map<String, Object> excelMap = new HashMap<>();
//      excelMap.put("headerKeys", headerKeys); // 헤더 정보
//      excelMap.put("widths", widths); // 칼럼 너비
//      excelMap.put("headerKeysMap", headerKeysMap); // 헤더에 따른 정보 매핑(리스트 사이즈만큼 로우 나옴)
//      excelMap.put("rowIndex", rowIndex);
//      excelMap.put("fileName", fileName);
//      // 10000개씩 페이징 처리 한 것을 엑셀로 변환
//      sxssfWorkbook = excelUtil.buildExcelDocument(excelMap, sxssfWorkbook, request, response);
//      headerKeysMap.clear(); //초기화
//    }
//    // 엑셀로 변환한 것을 파일로(FileOutputStream) 전환
//    excelUtil.writeSXSSFWorkbook(sxssfWorkbook, fileName, request, response);
//    long end = System.currentTimeMillis();
//    long gap = end - start;
//    log.info("소요시간 >> {}ms", gap);
//  }

  public void getExcelList(HttpServletRequest request, HttpServletResponse response) throws Exception {
    long start = System.currentTimeMillis();
    int page = 0;
    int size = 10000;
    Page<ProductDTO> productDTOS = getProductList(PageRequest.of(page, size));

    // 헤더당 셀 너비 설정
    List<String> widths = Arrays.asList("10", "20", "50", "15", "20");
    // 파일명 설정
    String fileName = "EXAMPLE_EXCEL";
    SXSSFWorkbook sxssfWorkbook = null;

    int totalPages = productDTOS.getTotalPages();
    boolean flag = false;
    for (int i = 0; i < totalPages; i++) {
      Page<ProductDTO> pagedExcelData = getProductList(PageRequest.of(i, size));

      List<Map<String, Object>> headerKeysMap = new ArrayList<>();
      for(ProductDTO productDTO : pagedExcelData){
        headerKeysMap.add(objectMapper.convertValue(productDTO, Map.class));
      }
      List<String> headerKeys = new ArrayList<>();
      headerKeys = excelHandler.getKeys(headerKeysMap.get(0), headerKeys);
      flag = pagedExcelData.isLast();
      int rowIndex = i * size;
      excelHandler.excelDownload(headerKeys, widths, headerKeysMap, rowIndex, flag, request, response);
//      Map<String, Object> excelMap = new HashMap<>();
//      excelMap.put("headerKeys", headerKeys); // 헤더 정보
//      excelMap.put("widths", widths); // 칼럼 너비
//      excelMap.put("headerKeysMap", headerKeysMap); // 헤더에 따른 정보 매핑(리스트 사이즈만큼 로우 나옴)
//      excelMap.put("rowIndex", rowIndex);
      // 10000개씩 페이징 처리 한 것을 엑셀로 변환
//      sxssfWorkbook = excelUtil.buildExcelDocument(excelMap, sxssfWorkbook);
      headerKeysMap.clear(); //초기화
    }
//    // 엑셀로 변환한 것을 파일로(FileOutputStream) 전환
//    excelUtil.writeSXSSFWorkbook(sxssfWorkbook, fileName, request, response);
    long end = System.currentTimeMillis();
    long gap = end - start;
    log.info("소요시간 >> {}ms", gap);
  }

  public ProductDTO createProduct(ProductDTO productDTO) {
    return Optional.of(Product.create(
        productDTO.getName(),
        productDTO.getDescription(),
        productDTO.getPrice()))
      .map(productRepository::save)
      .map(productMapper::toDTO)
      .get();
  }

  public Page<ProductDTO> getProductList(Pageable pageable) {
    return productRepository.findAll(pageable).map(productMapper::toDTO);
  }
}