package com.inventory.controller;

import com.inventory.domain.Outbound;
import com.inventory.domain.OutboundItem;
import com.inventory.service.InboundService;
import com.inventory.service.OutboundService;
import lombok.RequiredArgsConstructor;
import org.springframework.format.annotation.DateTimeFormat;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.time.LocalDate;
import java.util.List;
import java.util.stream.Collectors;

@Controller
@RequestMapping("/outbound")
@RequiredArgsConstructor
public class OutboundController {

    private final OutboundService outboundService;
    private final InboundService  inboundService;

    @GetMapping("/list")
    public String list(@RequestParam(required = false) Boolean useDate,
                       @RequestParam(required = false)
                           @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) LocalDate startDate,
                       @RequestParam(required = false)
                           @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) LocalDate endDate,
                       Model model) {
        boolean useDateFilter = Boolean.TRUE.equals(useDate);
        LocalDate effStart = useDateFilter ? startDate : null;
        LocalDate effEnd   = useDateFilter ? endDate   : null;

        model.addAttribute("outbounds", outboundService.findAll(effStart, effEnd));
        model.addAttribute("useDate",   useDateFilter);
        model.addAttribute("startDate", startDate);
        model.addAttribute("endDate",   endDate);
        model.addAttribute("today",     LocalDate.now().toString());
        return "outbound/list";
    }

    @GetMapping("/form")
    public String form(Model model) {
        // 잔여재고 > 0 인 입고만 보여줌 (전체에서 필터)
        model.addAttribute("availableInbounds",
                inboundService.findAll(null, null).stream()
                        .filter(i -> i.getRemainingQty() != null && i.getRemainingQty() > 0)
                        .collect(Collectors.toList()));
        model.addAttribute("today", LocalDate.now().toString());
        return "outbound/form";
    }

    /** AJAX - 출고 디테일 (모달용 JSON) */
    @GetMapping("/{id}/items")
    @ResponseBody
    public List<OutboundItem> items(@PathVariable String id) {
        return outboundService.findItemsByOutboundId(id);
    }

    @PostMapping("/save")
    public String save(@ModelAttribute Outbound outbound,
                       @RequestParam(required = false) List<String>  inboundIds,
                       @RequestParam(required = false) List<Integer> quantities,
                       RedirectAttributes ra) {
        try {
            outboundService.register(outbound, inboundIds, quantities);
            ra.addFlashAttribute("msg", "출고 [" + outbound.getOutboundId() + "] 등록 완료");
            return "redirect:/outbound/list";
        } catch (IllegalArgumentException | IllegalStateException e) {
            ra.addFlashAttribute("error", e.getMessage());
            return "redirect:/outbound/form";
        }
    }

    @PostMapping("/delete")
    public String delete(@RequestParam String outboundId, RedirectAttributes ra) {
        try {
            outboundService.delete(outboundId);
            ra.addFlashAttribute("msg", "출고 [" + outboundId + "] 삭제 완료. 입고 출고량이 원복되었습니다.");
        } catch (Exception e) {
            ra.addFlashAttribute("error", "삭제 중 오류: " + e.getMessage());
        }
        return "redirect:/outbound/list";
    }
}
