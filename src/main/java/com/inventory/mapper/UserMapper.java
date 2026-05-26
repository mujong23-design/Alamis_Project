package com.inventory.mapper;

import com.inventory.domain.User;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;

@Mapper
public interface UserMapper {
    User findById(@Param("userId") String userId);
    List<User> findAll();
    int insert(User user);
    int update(User user);                              // 이름, 권한 수정
    int updatePassword(@Param("userId") String userId,
                       @Param("userPw") String userPw); // 비밀번호만 별도 변경
    int delete(@Param("userId") String userId);
}
