package com.example.excel.test;

import com.example.excel.annotation.Excel;
import lombok.Data;

import java.util.Date;

/**
 * @Author: Jax
 * @Email: guoxingyou@xjiye.com
 * @Date: 2017/12/18/10:18
 * @Desc:
 **/
@Data
public class UserEntity {

    @Excel(title = "账户名", index = 0)
    private String name;

    @Excel(title = "密码", index = 0)
    private String password;

    @Excel(title = "年龄", index = 0, fieldType = int.class)
    private int age;

    @Excel(title = "UID", index = 0)
    private String uid;

    @Excel(title = "注册时间", index = 0, dateFmt = "yyyy-MM-dd HH:mm:ss")
    private Date regDate;


    public static class Builder{
        private String name;

        private String password;

        private int age;

        private String uid;

        private Date regDate;

        public Builder name(String name) {
            this.name = name;
            return this;
        }

        public Builder password(String password) {
            this.password = password;
            return this;
        }

        public Builder age(int age) {
            this.age = age;
            return this;
        }

        public Builder uid(String uid) {
            this.uid = uid;
            return this;
        }

        public Builder regDate(Date regDate) {
            this.regDate = regDate;
            return this;
        }

        public UserEntity build(){
            return new UserEntity(this);
        }
    }


    public UserEntity(Builder builder) {
        this.name = builder.name;
        this.password = builder.password;
        this.age = builder.age;
        this.uid = builder.uid;
        this.regDate = builder.regDate;
    }

    public UserEntity() {
    }
}
