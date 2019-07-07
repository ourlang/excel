package com.spring.excel.entity;

/**
 * 用户表
 * @author 福小林
 */
public class User {
    /**
     * 姓名
     */
    private String name;
    /**
     * 年龄
     */
    private String age;
    /**
     * 居住地址
     */
    private String address;
    /**
     * 邮箱
     */
    private String email;
    /**
     * 电话号码
     */
    private String telephone;

    public User(String name, String age, String address, String email, String telephone) {
        this.name = name;
        this.age = age;
        this.address = address;
        this.email = email;
        this.telephone = telephone;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getTelephone() {
        return telephone;
    }

    public void setTelephone(String telephone) {
        this.telephone = telephone;
    }


}
