package com.inventory.service;

import com.inventory.domain.Inbound;
import com.inventory.domain.Outbound;
import com.inventory.domain.OutboundItem;
import com.inventory.mapper.InboundMapper;
import com.inventory.mapper.OutboundItemMapper;
import com.inventory.mapper.OutboundMapper;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

@Service
@RequiredArgsConstructor
public class OutboundService {

    private static final DateTimeFormatter YYYYMMDD = DateTimeFormatter.ofPattern("yyyyMMdd");

    private final OutboundMapper      outboundMapper;
    private final OutboundItemMapper  outboundItemMapper;
    private final InboundMapper       inboundMapper;

    /* ===== 조회 ===== */
    public List<Outbound> findAll(LocalDate startDate, LocalDate endDate) {
        return outboundMapper.findAll(startDate, endDate);
    }

    public Outbound findById(String id) {
        Outbound outbound = outboundMapper.findById(id);
        if (outbound != null) {
            outbound.setItems(outboundItemMapper.findByOutboundId(id));
        }
        return outbound;
    }

    public List<OutboundItem> findItemsByOutboundId(String outboundId) {
        return outboundItemMapper.findByOutboundId(outboundId);
    }

    /* ===== 등록 ===== */

    /**
     * 출고 등록 — 트랜잭션:
     *  1) COA 필수 검증
     *  2) 각 입고 라인 잔여재고 검증
     *  3) outbound 마스터 INSERT (출고번호 자동채번)
     *  4) outbound_item 디테일 N개 INSERT
     *  5) inbound.outbound_qty 증가
     *  → 하나라도 실패 시 전부 롤백
     */
    @Transactional
    public void register(Outbound outbound, List<String> inboundIds, List<Integer> quantities) {
        // 1) 기본 검증
        if (outbound.getCoaCode() == null || outbound.getCoaCode().isBlank()) {
            throw new IllegalArgumentException("COA(성적서 코드)는 필수입니다.");
        }
        if (outbound.getOutboundDate() == null) {
            outbound.setOutboundDate(LocalDate.now());
        }
        if (inboundIds == null || inboundIds.isEmpty()) {
            throw new IllegalArgumentException("출고할 입고 제품을 1개 이상 선택해주세요.");
        }
        int n = Math.min(inboundIds.size(), quantities == null ? 0 : quantities.size());
        if (n == 0) {
            throw new IllegalArgumentException("출고 수량이 입력되지 않았습니다.");
        }

        // 2) 잔여재고 검증
        for (int i = 0; i < n; i++) {
            String inboundId = inboundIds.get(i);
            Integer qty = quantities.get(i);
            if (inboundId == null || inboundId.isBlank()) continue;
            if (qty == null || qty <= 0) {
                throw new IllegalArgumentException(inboundId + " 의 출고 수량은 1 이상이어야 합니다.");
            }
            Inbound inb = inboundMapper.findById(inboundId);
            if (inb == null) {
                throw new IllegalArgumentException("입고 정보 없음: " + inboundId);
            }
            int remaining = inb.getPackageCount() - inb.getOutboundQty();
            if (qty > remaining) {
                throw new IllegalStateException(
                        inboundId + " 잔여재고 부족 — 잔여: " + remaining + ", 요청: " + qty);
            }
        }

        // 3) 출고번호 자동채번 + 마스터 INSERT
        outbound.setOutboundId(generateNextId(outbound.getOutboundDate()));
        outboundMapper.insert(outbound);

        // 4) 디테일 INSERT + 5) 입고 출고량 증가
        for (int i = 0; i < n; i++) {
            String inboundId = inboundIds.get(i);
            Integer qty = quantities.get(i);
            if (inboundId == null || inboundId.isBlank() || qty == null || qty <= 0) continue;

            OutboundItem item = new OutboundItem();
            item.setOutboundId(outbound.getOutboundId());
            item.setInboundId(inboundId);
            item.setQuantity(qty);
            outboundItemMapper.insert(item);

            inboundMapper.incrementOutboundQty(inboundId, qty);
        }
    }

    /**
     * 출고 삭제 — 트랜잭션:
     *  1) 디테일 조회
     *  2) 각 입고 라인 outbound_qty 감소
     *  3) outbound 삭제 (CASCADE 로 디테일도 같이 삭제)
     */
    @Transactional
    public void delete(String outboundId) {
        List<OutboundItem> items = outboundItemMapper.findByOutboundId(outboundId);
        for (OutboundItem item : items) {
            inboundMapper.decrementOutboundQty(item.getInboundId(), item.getQuantity());
        }
        outboundMapper.delete(outboundId);
    }

    /** O-YYYYMMDD-NNN 형식 자동채번 */
    private String generateNextId(LocalDate date) {
        String dateStr = date.format(YYYYMMDD);
        Integer maxSeq = outboundMapper.findMaxSeqByDate(dateStr);
        int next = (maxSeq == null ? 0 : maxSeq) + 1;
        return String.format("O-%s-%03d", dateStr, next);
    }
}
