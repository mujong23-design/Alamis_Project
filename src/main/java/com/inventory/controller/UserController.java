package com.inventory.controller;

import com.inventory.domain.User;
import com.inventory.service.UserService;
import jakarta.servlet.http.HttpSession;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

@Controller
@RequestMapping("/user")
@RequiredArgsConstructor
public class UserController {

    private final UserService userService;

    /* ============ 관리자 전용 (AdminInterceptor 가 권한 체크) ============ */

    @GetMapping("/list")
    public String list(Model model) {
        model.addAttribute("users", userService.findAll());
        return "user/list";
    }

    @GetMapping("/form")
    public String form(@RequestParam(required = false) String userId, Model model) {
        boolean isEdit = userId != null;
        User user = isEdit ? userService.findById(userId) : new User();
        model.addAttribute("user", user);
        model.addAttribute("isEdit", isEdit);
        return "user/form";
    }

    @PostMapping("/save")
    public String save(@ModelAttribute User user,
                       @RequestParam(required = false) String rawPw,
                       @RequestParam(required = false, defaultValue = "false") boolean isEdit,
                       RedirectAttributes ra) {
        if (isEdit) {
            userService.update(user);
            if (rawPw != null && !rawPw.isBlank()) {
                userService.resetPassword(user.getUserId(), rawPw);
            }
            ra.addFlashAttribute("msg", "사용자 정보가 수정되었습니다.");
        } else {
            if (rawPw == null || rawPw.isBlank()) {
                ra.addFlashAttribute("msg", "신규 등록 시 비밀번호는 필수입니다.");
                return "redirect:/user/form";
            }
            userService.register(user, rawPw);
            ra.addFlashAttribute("msg", "사용자가 등록되었습니다.");
        }
        return "redirect:/user/list";
    }

    @PostMapping("/delete")
    public String delete(@RequestParam String userId,
                         HttpSession session,
                         RedirectAttributes ra) {
        User loginUser = (User) session.getAttribute("loginUser");
        if (loginUser != null && loginUser.getUserId().equals(userId)) {
            ra.addFlashAttribute("msg", "본인 계정은 삭제할 수 없습니다.");
            return "redirect:/user/list";
        }
        userService.delete(userId);
        ra.addFlashAttribute("msg", "사용자가 삭제되었습니다.");
        return "redirect:/user/list";
    }

    /* ============ 본인 비밀번호 변경 (모든 로그인 사용자) ============ */

    @GetMapping("/password")
    public String passwordForm() {
        return "user/password";
    }

    @PostMapping("/password")
    public String changePassword(@RequestParam String currentPw,
                                 @RequestParam String newPw,
                                 @RequestParam String newPwConfirm,
                                 HttpSession session,
                                 Model model) {
        User loginUser = (User) session.getAttribute("loginUser");
        if (loginUser == null) return "redirect:/login";

        if (!newPw.equals(newPwConfirm)) {
            model.addAttribute("error", "새 비밀번호 확인이 일치하지 않습니다.");
            return "user/password";
        }
        if (newPw.length() < 4) {
            model.addAttribute("error", "비밀번호는 4자 이상이어야 합니다.");
            return "user/password";
        }

        boolean ok = userService.changePassword(loginUser.getUserId(), currentPw, newPw);
        if (!ok) {
            model.addAttribute("error", "현재 비밀번호가 올바르지 않습니다.");
            return "user/password";
        }
        model.addAttribute("msg", "비밀번호가 변경되었습니다.");
        return "user/password";
    }
}
