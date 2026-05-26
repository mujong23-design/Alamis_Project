package com.inventory.controller;

import com.inventory.mapper.InboundMapper;
import com.inventory.mapper.OutboundMapper;
import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

@RestController
@RequestMapping("/dashboard")
@RequiredArgsConstructor
public class DashboardController {

    private final InboundMapper  inboundMapper;
    private final OutboundMapper outboundMapper;

    @GetMapping("/charts")
    public Map<String, Object> charts() {
        Map<String, Object> result = new HashMap<>();
        result.put("monthlyData", buildMonthlyData());
        result.put("topStocks",   inboundMapper.findTopStocks());
        return result;
    }

    /** 최근 6개월: 월별 입고량(SUM 포장개수) vs 출고량(SUM quantity), 빈 월은 0으로 채움 */
    private List<Map<String, Object>> buildMonthlyData() {
        DateTimeFormatter fmt = DateTimeFormatter.ofPattern("yyyy-MM");
        LocalDate today = LocalDate.now();

        Map<String, Integer> inMap = inboundMapper.findMonthlyInboundQty().stream()
                .collect(Collectors.toMap(
                        m -> (String) m.get("month"),
                        m -> ((Number) m.get("qty")).intValue()));

        Map<String, Integer> outMap = outboundMapper.findMonthlyOutboundQty().stream()
                .collect(Collectors.toMap(
                        m -> (String) m.get("month"),
                        m -> ((Number) m.get("qty")).intValue()));

        List<Map<String, Object>> result = new ArrayList<>();
        for (int i = 5; i >= 0; i--) {
            String month = today.minusMonths(i).format(fmt);
            Map<String, Object> entry = new HashMap<>();
            entry.put("month",    month);
            entry.put("inbound",  inMap.getOrDefault(month, 0));
            entry.put("outbound", outMap.getOrDefault(month, 0));
            result.add(entry);
        }
        return result;
    }
}
