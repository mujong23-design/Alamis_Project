package com.inventory.domain;

import lombok.Data;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;

/**
 * 입고 — 1건 = 1 row.
 * 출고는 outbound_qty 컬럼을 누적 증가시키는 방식으로 관리.
 * 현재 재고량 = package_count - outbound_qty
 */
@Data
public class Inbound {
    private String inboundId;            // PK: P-YYYYMMDD-NNN 자동채번
    private LocalDate productionDate;    // 생산일자
    private BigDecimal packageUnitKg;    // 포장 단위 (kg)
    private Integer packageCount;        // 포장 개수
    private BigDecimal sizeUm;           // 크기 (μm)
    private Integer outboundQty;         // 누적 출고량
    private String printedYn;            // 라벨 출력 이력 (Y/N)
    private String remark;
    private LocalDateTime createdAt;

    // 계산 필드 (DB 또는 Java 에서 계산)
    private Integer remainingQty;        // 잔여 재고 = packageCount - outboundQty
}
