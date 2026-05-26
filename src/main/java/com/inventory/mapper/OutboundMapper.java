package com.inventory.mapper;

import com.inventory.domain.Outbound;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.time.LocalDate;
import java.util.List;

@Mapper
public interface OutboundMapper {

    List<Outbound> findAll(@Param("startDate") LocalDate startDate,
                           @Param("endDate")   LocalDate endDate);

    Outbound findById(@Param("outboundId") String outboundId);

    int insert(Outbound outbound);

    int delete(@Param("outboundId") String outboundId);

    /** O-YYYYMMDD-NNN 형식 마지막 순번 */
    Integer findMaxSeqByDate(@Param("dateStr") String yyyymmdd);

    /* ===== 대시보드 통계 ===== */
    int countTodayOutbound();
    /** 금일 출고 총 수량 (오늘자 outbound 의 item quantity 합) */
    Long sumTodayOutboundQty();
    /** 최근 6개월 출고량 (월별 SUM(quantity)) */
    java.util.List<java.util.Map<String,Object>> findMonthlyOutboundQty();
}
