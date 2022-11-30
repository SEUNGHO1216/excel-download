package com.example.exceldownload.controller;

import com.example.exceldownload.dto.ProductDTO;
import com.example.exceldownload.repository.ProductRepository;
import com.example.exceldownload.service.ProductService;
import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

@RestController
@RequiredArgsConstructor
@CrossOrigin(origins = "*", allowedHeaders = "*")
public class ProductController {

  private final ProductService productService;
  private final ProductRepository productRepository;

  @GetMapping(value = "/excel/download", produces = "application/vnd.ms-excel")
  public String excelDownload(HttpServletRequest request, HttpServletResponse response) throws Exception {
    productService.getExcelList(request, response);
    return "엑셀 다운로드 완료";
  }

  @PostMapping("/product")
  public ProductDTO createProduct(@RequestBody ProductDTO productDTO){
    return productService.createProduct(productDTO);
  }

  @PostMapping("/sample")
  public Integer createSampleData(@RequestBody ProductDTO productDTO){
    for(int i = 0; i<10000; i++){
      productService.createProduct(productDTO);
    }
    return productRepository.findAll().size();
  }
}
