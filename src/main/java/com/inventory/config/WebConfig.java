package com.inventory.config;

import com.inventory.interceptor.AdminInterceptor;
import com.inventory.interceptor.LoginInterceptor;
import lombok.RequiredArgsConstructor;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.InterceptorRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

import java.util.List;

@Configuration
@RequiredArgsConstructor
public class WebConfig implements WebMvcConfigurer {

    private final LoginInterceptor loginInterceptor;
    private final AdminInterceptor adminInterceptor;

    @Override
    public void addInterceptors(InterceptorRegistry registry) {
        // 1) 로그인 체크 — /login, /logout, 정적자원만 제외
        registry.addInterceptor(loginInterceptor)
                .order(1)
                .addPathPatterns("/**")
                .excludePathPatterns(
                        "/login", "/logout",
                        "/css/**", "/js/**", "/images/**", "/error"
                );

        // 2) 관리자 권한 체크 — 사용자 관리 화면만 (본인 비번 변경은 제외)
        registry.addInterceptor(adminInterceptor)
                .order(2)
                .addPathPatterns(List.of(
                        "/user/list",
                        "/user/form",
                        "/user/save",
                        "/user/delete"
                ));
    }
}
