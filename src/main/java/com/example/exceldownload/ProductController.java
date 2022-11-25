package com.example.exceldownload;

import lombok.RequiredArgsConstructor;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequiredArgsConstructor
public class ProductController {

  private final ProductService productService;

  @GetMapping(value = "/excel/download", produces = "application/vnd.ms-excel")
  public String excelDownload() {
    productService.getExcelList();
    return "엑셀 다운로드 완료";
  }
}
