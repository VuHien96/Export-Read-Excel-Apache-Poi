package com.vuhien.application.service;

import com.vuhien.application.entity.Nation;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.List;

@Component
public class NationServiceImpl implements NationService{

    private static ArrayList<Nation> nations = new ArrayList<>();

    static{
        nations.add(new Nation(1,"Russia",100));
        nations.add(new Nation(2,"Canada",400));
        nations.add(new Nation(3,"Brazil",200));
        nations.add(new Nation(4,"China",100));
        nations.add(new Nation(5,"United States",600));
    }

    @Override
    public List<Nation> getListNations() {
        return nations;
    }
}
