package com.inventory.controller;

import com.inventory.domain.Inbound;
import com.inventory.service.InboundService;
import lombok.RequiredArgsConstructor;
import org.springframework.format.annotation.DateTimeFormat;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.time.LocalDate;

@Controller
@RequestMapping("/inbound")
@RequiredArgsConstructor
public class InboundController {

    private final InboundService inboundService;

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

        model.addAttribute("inbounds",  inboundService.findAll(effStart, effEnd));
        model.addAttribute("useDate",   useDateFilter);
        model.addAttribute("startDate", startDate);
        model.addAttribute("endDate",   endDate);
        model.addAttribute("today",     LocalDate.now().toString());
        return "inbound/list";
    }

    @GetMapping("/form")
    public String form(@RequestParam(required = false) String id, Model model) {
        boolean isEdit = (id != null);
        Inbound inbound = isEdit ? inboundService.findById(id) : new Inbound();
        model.addAttribute("inbound", inbound);
        model.addAttribute("isEdit",  isEdit);
        model.addAttribute("today",   LocalDate.now().toString());
        return "inbound/form";
    }

    @PostMapping("/save")
    public String save(@ModelAttribute Inbound inbound, RedirectAttributes ra) {
        try {
            if (inbound.getInboundId() == null || inbound.getInboundId().isBlank()) {
                inboundService.register(inbound);
                ra.addFlashAttribute("msg", "입고가 등록되었습니다.");
            } else {
                inboundService.update(inbound);
                ra.addFlashAttribute("msg", "입고 정보가 수정되었습니다.");
            }
        } catch (Exception e) {
            ra.addFlashAttribute("error", "저장 중 오류: " + e.getMessage());
            return "redirect:/inbound/form";
        }
        return "redirect:/inbound/list";
    }

    @PostMapping("/delete")
    public String delete(@RequestParam String inboundId, RedirectAttributes ra) {
        try {
            inboundService.delete(inboundId);
            ra.addFlashAttribute("msg", "입고가 삭제되었습니다.");
        } catch (Exception e) {
            ra.addFlashAttribute("error", "삭제 중 오류: " + e.getMessage());
        }
        return "redirect:/inbound/list";
    }

    /** 포장 라벨 인쇄 페이지 (선택된 입고건 N개 → 각 포장개수만큼 라벨) */
    @GetMapping("/label")
    public String label(@RequestParam("ids") java.util.List<String> ids,
                        org.springframework.ui.Model model) {
        model.addAttribute("inbounds", inboundService.findByIds(ids));
        // 라벨 페이지 진입 = 출력으로 간주, 출력여부 'Y' 갱신
        inboundService.markPrinted(ids);
        return "inbound/label";
    }
}
