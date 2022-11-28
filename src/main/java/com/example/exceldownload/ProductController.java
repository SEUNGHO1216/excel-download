package com.example.exceldownload;

import lombok.RequiredArgsConstructor;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

@RestController
@RequiredArgsConstructor
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
