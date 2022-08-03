package com.hongke.easyexcel;

import org.junit.Test;

import java.util.Optional;

public class OptionalTest {

    @Test
    public void test() {
        DemoData demoData = new DemoData();
        demoData.setString("2335");
        String s = Optional.ofNullable(demoData).map(x -> demoData.getString()).orElse("123");
        System.out.println(s);
    }

}
