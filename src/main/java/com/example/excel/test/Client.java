package com.example.excel.test;

import com.example.excel.enums.OperationType;
import com.example.excel.handle.ExportHandle;
import com.example.excel.handle.ImportHandle;
import com.google.common.collect.Lists;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.List;

/**
 * @Author: Jax
 * @Email: guoxingyou@xjiye.com
 * @Date: 2017/12/18/10:13
 * @Desc:
 **/
public class Client {


    public static void main(String[] args) {
        List<UserEntity> list = Lists.newArrayList();

        UserEntity userEntity = new UserEntity.Builder().name("A").password("123").age(12).uid
                ("456").regDate(new Date()).build();

        UserEntity userEntity1 = new UserEntity.Builder().name("B").password("234").age(12).uid
                ("567").regDate(new Date()).build();

        UserEntity userEntity2 = new UserEntity.Builder().name("C").password("345").age(12).uid
                ("678").regDate(new Date()).build();

        list.add(userEntity);list.add(userEntity1);list.add(userEntity2);

        try {
            new ExportHandle("test",UserEntity.class, OperationType.ONLY_EXPORT).setDataList(list)
                    .writeFile("target/export.xlsx").dispose();
        } catch (IOException e) {
            e.printStackTrace();
        }

        try {
            List<UserEntity> list2 =  new ImportHandle(new File("target/export.xlsx"),1).getDataList(UserEntity.class);
            System.out.println(list2.toString());
        } catch (Exception e) {
            e.printStackTrace();
        }


    }

}
