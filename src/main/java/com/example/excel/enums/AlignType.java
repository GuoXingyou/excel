package com.example.excel.enums;

/**
 * @Author: Jax
 * @Date: 2017-12-2017/12/14-17:41
 * @Desc:
 **/
public enum AlignType {

    CENTER("CENTER",2,"居中"),

    RIGHT("RIGHT",3,"靠右"),

    LEFT("LEFT",1,"靠左"),

    AUTO("AUTO",0,"自动")

    ;

    private String code;

    private int num;

    private String message;

    private AlignType(String code, int num, String message) {
        this.code = code;
        this.num = num;
        this.message = message;
    }

    public String code() {
        return code;
    }

    public int num() { return num; }

    public String message() {
        return message;
    }

    public static AlignType getByCode(String code) {
        for (AlignType result : values()) {
            if (result.code().equals(code)) {
                return result;
            }
        }
        return null;
    }

    /**
     * 获取全部枚举值
     *
     * @return List<String>
     */
    public static java.util.List<String> getAllEnumCode() {
        java.util.List<String> list = new java.util.ArrayList<String>(values().length);
        for (AlignType _enum : values()) {
            list.add(_enum.code());
        }
        return list;
    }

    /**
     * 获取全部枚举
     *
     * @return List<ListType>
     */
    public static java.util.List<AlignType> getAllEnum() {
        java.util.List<AlignType> list = new java.util.ArrayList<AlignType>(values().length);
        for (AlignType _enum : values()) {
            list.add(_enum);
        }
        return list;
    }

    /**
     * 获取全部枚举描述
     *
     * @return List<ListType>
     */
    public static java.util.List<String> getAllEnumMessage() {
        java.util.List<String> list = new java.util.ArrayList<String>(values().length);
        for (AlignType _enum : values()) {
            list.add(_enum.message);
        }
        return list;
    }

}
