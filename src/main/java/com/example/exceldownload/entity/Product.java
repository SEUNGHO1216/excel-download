package com.example.exceldownload.entity;

import lombok.*;

import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import java.time.LocalDateTime;

@Entity
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class Product {

  @Id
  @GeneratedValue(strategy = GenerationType.IDENTITY)
  private Long id;
  private String name;
  private String description;
  private Integer price;
  private LocalDateTime expireDate;

  public static Product create(String name,
                               String description,
                               Integer price){
    Product product = new Product();
    product.setName(name);
    product.setDescription(description);
    product.setPrice(price);
    product.setExpireDate(LocalDateTime.now().plusMonths(1L));
    return product;
  }
}
