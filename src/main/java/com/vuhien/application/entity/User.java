package com.vuhien.application.entity;

import lombok.*;

import java.util.Date;

@AllArgsConstructor
@NoArgsConstructor
@Setter
@Getter
@ToString
public class User {
    private long id;
    private String email;
    private String password;
    private String fullName;
    private long age;
    private String address;
    private String phone;
    private Date createdAt;
    private Date modifiedAt;
}
