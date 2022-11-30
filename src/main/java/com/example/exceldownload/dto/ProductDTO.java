package com.example.exceldownload.dto;

import lombok.Data;

import java.time.LocalDateTime;

@Data
public class ProductDTO {
  private Long id;
  private String name;
  private String description;
  private Integer price;
  private LocalDateTime expireDate;
}
