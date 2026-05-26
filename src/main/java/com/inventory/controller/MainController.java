package com.inventory.controller;

import com.inventory.mapper.InboundMapper;
import com.inventory.mapper.OutboundMapper;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;

@Controller
@RequiredArgsConstructor
public class MainController {

    private final InboundMapper  inboundMapper;
    private final OutboundMapper outboundMapper;

    @GetMapping("/")
    public String root() {
        return "redirect:/main";
    }

    @GetMapping("/main")
    public String main(Model model) {
        model.addAttribute("totalInbounds",   inboundMapper.countAll());
        model.addAttribute("currentStock",    inboundMapper.sumRemainingStock());
        model.addAttribute("todayInbounds",   inboundMapper.countTodayInbound());
        model.addAttribute("todayOutbounds",  outboundMapper.countTodayOutbound());
        model.addAttribute("todayOutboundQty", outboundMapper.sumTodayOutboundQty());
        return "main";
    }
}
