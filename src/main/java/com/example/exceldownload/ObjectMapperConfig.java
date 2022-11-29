package com.example.exceldownload;

import com.fasterxml.jackson.annotation.JsonTypeInfo;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class ObjectMapperConfig {

  @Bean
  public ObjectMapper objectMapper() {

//    ObjectMapper objectMapper = new ObjectMapper();
//    objectMapper.activateDefaultTyping(
//      objectMapper.getPolymorphicTypeValidator(),
//      ObjectMapper.DefaultTyping.NON_FINAL,
//      JsonTypeInfo.As.PROPERTY
//    )
//    .findAndRegisterModules()
//    .enable(SerializationFeature.INDENT_OUTPUT)
//    .disable(SerializationFeature.WRITE_DATES_AS_TIMESTAMPS)
//    .configure(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES, false)
//    .registerModule(new JavaTimeModule());
//    return objectMapper;
    ObjectMapper objectMapper = new ObjectMapper();
    objectMapper.registerModule(new JavaTimeModule());
    objectMapper.disable(SerializationFeature.WRITE_DATES_AS_TIMESTAMPS);
    return objectMapper;
  }
}
