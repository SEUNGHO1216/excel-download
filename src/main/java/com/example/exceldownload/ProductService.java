package com.example.exceldownload;

import lombok.RequiredArgsConstructor;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Pageable;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.Field;
import java.util.*;

@Service
@RequiredArgsConstructor
public class ProductService {

  private final ProductRepository productRepository;
  private final ProductMapper productMapper;
  private final ExcelUtil excelUtil;

  public Map<String, Object> getExcelList(HttpServletRequest request, HttpServletResponse response) throws Exception {
    int page = 0;
    int size = 10000;
//    Page<ProductDTO> productDTOS = getProductList(PageRequest.of(page, size));
//
////    if(productDTOS.getSize() != 0){
//    Class<ProductDTO> productDTOClass = (Class<ProductDTO>) productDTOS.getContent().get(0).getClass();
//    Field[] productFields = productDTOClass.getFields();
//    int columnLength = productFields.length;
//    List<String> headerKeys = new ArrayList<>();
//    for (int i = 0; i < columnLength; i++) {
//      headerKeys.add(productFields[i].getName());
//    }
//    }

    // 엑셀에 저장할 row 수
//    int rowSize = productDTOS.getContent().size();
//    int totalPages = productDTOS.getTotalPages();
    /*
     // 엑셀 헤더 키 설정
      필드값을 자바 메소드를 활용해서 뽑는 방법은 아래를 참조
      https://roytuts.com/handling-large-data-writing-to-excel-using-sxssf-apache-poi/
     */
    List<String> headerKeys = List.of("id", "name", "description", "price", "expireDate"); // 이 자체가 헤더 키이자 헤더이다
    // 헤더당 셀 너비 설정
    List<String> widths = Arrays.asList("10", "20", "50", "15", "20");
//
//    List<Map<String, Object>> headerKeysMap = new ArrayList<>();
//    int totalPages = productDTOS.getTotalPages();
//    for (int i = 0; i < totalPages; i++) {
//      Page<ProductDTO> pagedExcelData = getProductList(PageRequest.of(i, size));
//      int rowIndex = i * page;
//      // 헤더 키에 1:1 매핑, 만개의 리스트 == 만개의 로우
//      for (ProductDTO excelData : pagedExcelData) {
//        Map<String, Object> tempMap = new HashMap<>();
//        for (Field field : productFields){
//          field.getName()
//        }
//          tempMap.put("id", excelData.getId());
//        tempMap.put("name", excelData.getName());
//        tempMap.put("description", excelData.getDescription());
//        tempMap.put("price", excelData.getPrice());
//        tempMap.put("expireDate", excelData.getExpireDate());
//
//        headerKeysMap.add(tempMap);
//      }
//    }


    // 파일명 설정
    String fileName = "EXAMPLE_EXCEL";

    Map<String, Object> excelMap = new HashMap<>();
    excelMap.put("headerKeys", headerKeys); // 헤더 정보
    excelMap.put("widths", widths); // 칼럼 너비
//    excelMap.put("headerKeysMap", headerKeysMap); // 헤더에 따른 정보 매핑(리스트 사이즈만큼 로우 나옴)
    excelMap.put("fileName", fileName);

    excelUtil.buildExcelDocument(excelMap, request, response);
    return excelMap;
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


/*
 * 흐름정리
 * 1. 페이징 된 데이터 넘기기(예를 들어 만개 씩 페이징/ 총 데이터는 3만5천)
 * 2. 파라미터로 시작값 넘기기(예를 들어 0, 10000, 20000, 30000 % max row -> 처음인지 아닌지 분간 및 rowNo 설정 영향)
 * 3. 만약 파라미터로 시작값이 아닌 페이지 번호가 들어오면 x 갯수를 해주면 시작값일 것이다. */
