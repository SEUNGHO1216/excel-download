package com.example.exceldownload.service;

import com.example.exceldownload.util.ExcelHandler;
import com.example.exceldownload.dto.ProductDTO;
import com.example.exceldownload.mapper.ProductMapper;
import com.example.exceldownload.repository.ProductRepository;
import com.example.exceldownload.entity.Product;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Pageable;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.*;

@Slf4j
@Service
@RequiredArgsConstructor
public class ProductService {

  private final ProductRepository productRepository;
  private final ProductMapper productMapper;
  private final ObjectMapper objectMapper;

  public void getExcelList(HttpServletRequest request, HttpServletResponse response) throws Exception {

    ExcelHandler excelHandler = new ExcelHandler();
    long start = System.currentTimeMillis();
    int page = 0;
    final int size = 10000;
    Page<ProductDTO> productDTOS = getProductList(PageRequest.of(page, size));

    // 헤더당 셀 너비 설정
    List<String> widths = Arrays.asList("10", "20", "50", "15", "20");

    int totalPages = productDTOS.getTotalPages();
    for (int i = 0; i < totalPages; i++) {
      Page<ProductDTO> pagedExcelData = getProductList(PageRequest.of(i, size));
      int rowIndex = i * size; //0, 10000, 20000 ... 페이징 처리를 통한 첫 페이지 명시

      List<Map<String, Object>> headerKeysMap = new ArrayList<>();
      for(ProductDTO excelData : pagedExcelData){
        // excelData 객체를 objectMapper 를 통해 key-value map 으로 전환
        headerKeysMap.add(objectMapper.convertValue(excelData, Map.class));

      }
      // map 에서 key 값을 헤더로 사용
      List<String> headerKeys = new ArrayList<>();
      headerKeys = excelHandler.getKeys(headerKeysMap.get(0), headerKeys);

      // isLast 가 true 일때 output stream 을 통해 파일을 쓴다
      boolean isLast = pagedExcelData.isLast();
      // excelHandler 의 다운로드 기능 call
      excelHandler.excelDownload(headerKeys, widths, headerKeysMap, rowIndex, isLast, request, response);

//      headerKeysMap.clear(); //초기화
    }
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