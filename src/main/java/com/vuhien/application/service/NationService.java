package com.vuhien.application.service;

import com.vuhien.application.entity.Nation;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public interface NationService {
    List<Nation> getListNations();
}
