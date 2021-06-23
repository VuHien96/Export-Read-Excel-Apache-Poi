package com.vuhien.application.service;

import com.vuhien.application.entity.User;
import com.vuhien.application.model.dto.UserDTO;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@Component
public class UserServiceImpl implements UserService {
    private static ArrayList<User> users = new ArrayList<>();

    static {
        users.add(new User(1, "a@gmail.com", "123456", "Nguyễn Văn A", 12, "Hà Nội", "0969708715", new Date(), null));
        users.add(new User(2, "b@gmail.com", "123456", "Nguyễn Văn B", 12, "Hà Nội", "0969708715", new Date(), null));
        users.add(new User(3, "c@gmail.com", "123456", "Nguyễn Văn C", 16, "Hà Nội", "0969708715", new Date(), null));
        users.add(new User(4, "d@gmail.com", "123456", "Nguyễn Văn D", 17, "Hà Nội", "0969708715", new Date(), null));
        users.add(new User(5, "e@gmail.com", "123456", "Nguyễn Văn E", 18, "Hà Nội", "0969708715", new Date(), null));
        users.add(new User(6, "f@gmail.com", "123456", "Nguyễn Văn F", 21, "Hà Nội", "0969708715", new Date(), null));
        users.add(new User(7, "g@gmail.com", "123456", "Nguyễn Văn G", 21, "Hà Nội", "0969708715", new Date(), null));
        users.add(new User(8, "h@gmail.com", "123456", "Nguyễn Văn H", 19, "Hà Nội", "0969708715", new Date(), null));
        users.add(new User(9, "j@gmail.com", "123456", "Nguyễn Văn J", 32, "Hà Nội", "0969708715", new Date(), null));
        users.add(new User(10, "k@gmail.com", "123456", "Nguyễn Văn K", 21, "Hà Nội", "0969708715", new Date(), null));
    }


    @Override
    public List<UserDTO> getListUsers() {
        List<UserDTO> userDTOS = new ArrayList<>();
        for (User user : users) {
            UserDTO userDTO = new UserDTO();
            userDTO.setId(user.getId());
            userDTO.setEmail(user.getEmail());
            userDTO.setFullName(user.getFullName());
            userDTO.setAge(user.getAge());
            userDTO.setPhone(user.getPhone());
            userDTO.setAddress(user.getAddress());
            userDTO.setCreatedAt(user.getCreatedAt());
            userDTO.setModifiedAt(user.getModifiedAt());
            userDTOS.add(userDTO);
        }
        return userDTOS;
    }
}
