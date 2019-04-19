package com.springboot.easyexcel.exception;

/**
 * 自定义解析excel的异常
 * @auther liupeiqing
 * @date 2019/4/19 13:22
 */
public class ExcelException extends RuntimeException{

    public ExcelException(String msg){
        super(msg);
    }
}
