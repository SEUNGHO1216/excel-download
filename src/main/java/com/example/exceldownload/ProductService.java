package com.example.exceldownload;

import lombok.RequiredArgsConstructor;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Map;

@Service
@RequiredArgsConstructor
public class ProductService {

  private final ProductRepository productRepository;
  private final ProductMapper productMapper;

  public Map<String, Object> getExcelList() {
    // 엑셀에 저장할 데이터
    List<ProductDTO> excelDataList =
      productMapper.toDTO(productRepository.findAll(Sort.by(Sort.Direction.DESC, "id")));
    // 엑셀에 저장할 row 수
    int dataSize = excelDataList.size();
    // 엑셀 헤더 키 설정
    List<String> headerKeys = List.of("id", "name", "description", "price", "expireDate");


  }
}
