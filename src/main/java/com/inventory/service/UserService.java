package com.inventory.service;

import com.inventory.domain.User;
import com.inventory.mapper.UserMapper;
import lombok.RequiredArgsConstructor;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;

@Service
@RequiredArgsConstructor
public class UserService {

    private final UserMapper userMapper;
    private final PasswordEncoder passwordEncoder;

    /** 로그인 — BCrypt 매칭 */
    public User login(String userId, String rawPw) {
        User user = userMapper.findById(userId);
        if (user == null) return null;
        if (!passwordEncoder.matches(rawPw, user.getUserPw())) return null;
        return user;
    }

    public List<User> findAll() {
        return userMapper.findAll();
    }

    public User findById(String userId) {
        return userMapper.findById(userId);
    }

    /** 관리자: 신규 사용자 등록 (비밀번호는 BCrypt 암호화) */
    @Transactional
    public void register(User user, String rawPw) {
        user.setUserPw(passwordEncoder.encode(rawPw));
        userMapper.insert(user);
    }

    /** 관리자: 사용자 정보 수정 (이름, 권한) */
    @Transactional
    public void update(User user) {
        userMapper.update(user);
    }

    /** 관리자: 비밀번호 초기화 */
    @Transactional
    public void resetPassword(String userId, String newRawPw) {
        userMapper.updatePassword(userId, passwordEncoder.encode(newRawPw));
    }

    /** 본인: 비밀번호 변경 (현재 비밀번호 검증 후) */
    @Transactional
    public boolean changePassword(String userId, String currentPw, String newPw) {
        User user = userMapper.findById(userId);
        if (user == null) return false;
        if (!passwordEncoder.matches(currentPw, user.getUserPw())) return false;
        userMapper.updatePassword(userId, passwordEncoder.encode(newPw));
        return true;
    }

    @Transactional
    public void delete(String userId) {
        userMapper.delete(userId);
    }
}
