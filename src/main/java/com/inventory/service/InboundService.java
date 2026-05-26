package com.inventory.service;

import com.inventory.domain.Inbound;
import com.inventory.mapper.InboundMapper;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

@Service
@RequiredArgsConstructor
public class InboundService {

    private static final DateTimeFormatter YYYYMMDD = DateTimeFormatter.ofPattern("yyyyMMdd");

    private final InboundMapper inboundMapper;

    public List<Inbound> findAll(LocalDate startDate, LocalDate endDate) {
        return inboundMapper.findAll(startDate, endDate);
    }

    public Inbound findById(String id) {
        return inboundMapper.findById(id);
    }

    public List<Inbound> findByIds(List<String> ids) {
        if (ids == null || ids.isEmpty()) return java.util.Collections.emptyList();
        return inboundMapper.findByIds(ids);
    }

    /** 선택된 입고건들의 라벨 출력 이력을 'Y' 로 갱신 */
    @Transactional
    public void markPrinted(List<String> ids) {
        if (ids != null && !ids.isEmpty()) {
            inboundMapper.updatePrintedYn(ids, "Y");
        }
    }

    /**
     * 입고 등록 — 제품번호 자동채번 (P-YYYYMMDD-NNN).
     * 생산일자 기준으로 NNN 부여.
     */
    @Transactional
    public void register(Inbound inbound) {
        if (inbound.getProductionDate() == null) {
            inbound.setProductionDate(LocalDate.now());
        }
        if (inbound.getOutboundQty() == null) {
            inbound.setOutboundQty(0);
        }
        inbound.setInboundId(generateNextId(inbound.getProductionDate()));
        inboundMapper.insert(inbound);
    }

    @Transactional
    public void update(Inbound inbound) {
        inboundMapper.update(inbound);
    }

    @Transactional
    public void delete(String id) {
        inboundMapper.delete(id);
    }

    /** P-YYYYMMDD-NNN 형식 자동 채번 (해당 일자의 max + 1) */
    private String generateNextId(LocalDate date) {
        String dateStr = date.format(YYYYMMDD);
        Integer maxSeq = inboundMapper.findMaxSeqByDate(dateStr);
        int next = (maxSeq == null ? 0 : maxSeq) + 1;
        return String.format("P-%s-%03d", dateStr, next);
    }
}
