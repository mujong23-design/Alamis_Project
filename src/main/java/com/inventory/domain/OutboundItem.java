package com.inventory.domain;

import lombok.Data;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;

/**
 * 출고 디테일 — 출고 1건에 들어가는 개별 입고 제품과 그 수량.
 */
@Data
public class OutboundItem {
    private Long itemId;
    private String outboundId;
    private String inboundId;
    private Integer quantity;
    private LocalDateTime createdAt;

    // 조회용 조인 필드 (inbound 마스터에서)
    private LocalDate productionDate;
    private BigDecimal packageUnitKg;
    private BigDecimal sizeUm;
}
