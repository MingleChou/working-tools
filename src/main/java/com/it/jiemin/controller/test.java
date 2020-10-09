package com.it.jiemin.controller;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * @author JieminZhou
 * @version 1.0
 * @date 2020/10/9 16:49
 */
@RestController
@RequestMapping("/")
public class test {
    @GetMapping("hello")
    public String testH(){
        return "Hello world!";
    }
}
