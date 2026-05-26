package com.inventory.mapper;

import com.inventory.domain.Inbound;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.time.LocalDate;
import java.util.List;

@Mapper
public interface InboundMapper {

    /** 조건 검색 (자재/일자 범위 — 추후 확장 대비) */
    List<Inbound> findAll(@Param("startDate") LocalDate startDate,
                          @Param("endDate")   LocalDate endDate);

    Inbound findById(@Param("inboundId") String inboundId);

    /** 여러 입고 일괄 조회 (라벨 인쇄용) */
    List<Inbound> findByIds(@Param("ids") List<String> ids);

    int insert(Inbound inbound);

    int update(Inbound inbound);

    int delete(@Param("inboundId") String inboundId);

    /** 같은 생산일자의 마지막 채번 (NNN 부분). 없으면 0 */
    Integer findMaxSeqByDate(@Param("dateStr") String yyyymmdd);

    /** 출고량 증가 (출고 등록 시 호출). CHECK 제약으로 package_count 초과 시 DB 에러 */
    int incrementOutboundQty(@Param("inboundId") String inboundId,
                             @Param("delta")     Integer delta);

    /** 출고량 감소 (출고 삭제 시 호출) */
    int decrementOutboundQty(@Param("inboundId") String inboundId,
                             @Param("delta")     Integer delta);

    /** 라벨 출력 여부 갱신 (여러 행 일괄) */
    int updatePrintedYn(@Param("ids")  List<String> ids,
                        @Param("flag") String       flag);

    /* ===== 대시보드 통계 ===== */
    int countAll();
    /** 현재 재고 (잔여 포장 수의 합) */
    Long sumRemainingStock();
    /** 금일 생산 건수 */
    int countTodayInbound();
    /** 최근 6개월 입고량 (월별 SUM(package_count)) */
    java.util.List<java.util.Map<String,Object>> findMonthlyInboundQty();
    /** 제품별 현재 재고 TOP 10 (remaining > 0) */
    java.util.List<java.util.Map<String,Object>> findTopStocks();
}
