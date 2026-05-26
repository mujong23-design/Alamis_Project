package com.inventory.domain;

import lombok.Data;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;

/**
 * 출고 마스터 — 1 출고건 = N개의 입고 제품을 묶어서 출고.
 */
@Data
public class Outbound {
    private String outboundId;            // O-YYYYMMDD-NNN
    private LocalDate outboundDate;
    private String coaCode;                // 성적서 코드 (필수)
    private String outboundLocation;
    private String remark;
    private LocalDateTime createdAt;

    // 집계용 (조회 시)
    private Integer totalQty;              // 디테일 quantity 합계
    private Integer itemCount;             // 디테일 행 수

    // 상세 조회용
    private List<OutboundItem> items;
}
