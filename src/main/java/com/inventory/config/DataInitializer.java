package com.inventory.config;

import com.inventory.domain.User;
import com.inventory.mapper.UserMapper;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.CommandLineRunner;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.stereotype.Component;

/**
 * 최초 실행 시 기본 관리자 계정이 없으면 생성한다.
 * - ID: admin / PW: admin1234 (BCrypt 암호화)
 */
@Slf4j
@Component
@RequiredArgsConstructor
public class DataInitializer implements CommandLineRunner {

    private static final String DEFAULT_ADMIN_ID  = "admin";
    private static final String DEFAULT_ADMIN_PW  = "admin1234";
    private static final String DEFAULT_ADMIN_NM  = "관리자";

    private final UserMapper userMapper;
    private final PasswordEncoder passwordEncoder;

    @Override
    public void run(String... args) {
        if (userMapper.findById(DEFAULT_ADMIN_ID) == null) {
            User admin = new User();
            admin.setUserId(DEFAULT_ADMIN_ID);
            admin.setUserPw(passwordEncoder.encode(DEFAULT_ADMIN_PW));
            admin.setUserName(DEFAULT_ADMIN_NM);
            admin.setRole("ADMIN");
            userMapper.insert(admin);
            log.info("==> 기본 관리자 계정 생성: {} / {}", DEFAULT_ADMIN_ID, DEFAULT_ADMIN_PW);
        }
    }
}
