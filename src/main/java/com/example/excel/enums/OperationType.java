package com.example.excel.enums;

/**
 * @Author: Jax
 * @Date: 2017-12-2017/12/14-17:44
 * @Desc:
 **/
public enum OperationType {

    ONLY_IMPORT("IMPORT",2,"仅导入"),

    ONLY_EXPORT("EXPORT",1,"仅导出"),

    BOTH("BOTH",0,"导入导出")

    ;

    private String code;

    private int num;

    private String message;

    private OperationType(String code, int num, String message) {
        this.code = code;
        this.num = num;
        this.message = message;
    }

    public String code() {
        return code;
    }

    public int num() { return  num; }

    public String message() {
        return message;
    }

    public static OperationType getByCode(String code) {
        for (OperationType result : values()) {
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
        for (OperationType _enum : values()) {
            list.add(_enum.code());
        }
        return list;
    }

    /**
     * 获取全部枚举
     *
     * @return List<ListType>
     */
    public static java.util.List<OperationType> getAllEnum() {
        java.util.List<OperationType> list = new java.util.ArrayList<OperationType>(values().length);
        for (OperationType _enum : values()) {
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
        for (OperationType _enum : values()) {
            list.add(_enum.message);
        }
        return list;
    }

}
