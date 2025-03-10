package com.api.excel.exporter.service;

import com.api.excel.exporter.dto.ProductDTO;
import com.api.excel.exporter.entity.Product;
import com.api.excel.exporter.repository.ProductRepository;
import lombok.RequiredArgsConstructor;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.Pageable;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
@RequiredArgsConstructor
public class ProductService {
    private final ProductRepository productRepository;

    public Page<ProductDTO> getProducts(Pageable pageable) {
        Page<ProductDTO> products =  productRepository.findAll(pageable)
                .map(product -> new ProductDTO(product.getId(), product.getName(), product.getCategory(), product.getPrice(), product.getDateAdded()));
        return products;
    }
}
