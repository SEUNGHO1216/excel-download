package com.example.exceldownload.mapper;

import com.example.exceldownload.dto.ProductDTO;
import com.example.exceldownload.entity.Product;
import org.mapstruct.Mapper;
import org.mapstruct.Mapping;

import java.util.List;

@Mapper(componentModel = "spring")
public interface ProductMapper {

  @Mapping(target = "id", source = "id")
  @Mapping(target = "name", source = "name")
  @Mapping(target = "description", source = "description")
  @Mapping(target = "price", source = "price")
  @Mapping(target = "expireDate", source = "expireDate")
  ProductDTO toDTO(Product product);

  List<ProductDTO> toDTO(List<Product> products);
}
