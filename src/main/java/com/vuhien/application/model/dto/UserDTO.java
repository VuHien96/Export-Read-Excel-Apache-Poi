package com.vuhien.application.model.dto;

import lombok.*;

import java.util.Date;

@AllArgsConstructor
@NoArgsConstructor
@Setter
@Getter
@ToString
public class UserDTO {
    private long id;
    private String email;
    private String fullName;
    private long age;
    private String address;
    private String phone;
    private Date createdAt;
    private Date modifiedAt;
}
