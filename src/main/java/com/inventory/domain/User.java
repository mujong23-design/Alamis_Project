package com.inventory.domain;

import lombok.Data;
import java.time.LocalDateTime;

@Data
public class User {
    private String userId;
    private String userPw;
    private String userName;
    private String role;
    private LocalDateTime createdAt;
}
