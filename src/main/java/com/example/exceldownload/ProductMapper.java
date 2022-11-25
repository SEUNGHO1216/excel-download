package com.example.exceldownload;

import org.mapstruct.Mapper;

import java.util.List;

@Mapper(componentModel = "spring")
public interface ProductMapper {

  ProductDTO toDTO(Product product);

  List<ProductDTO> toDTO(List<Product> products);
}
