package com.spring.excel.controller;

import com.spring.excel.entity.User;
import com.spring.excel.utils.ExcelUtils;
import org.apache.commons.collections4.map.LinkedMap;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

/**
 * 导出excel数据
 * @author 福小林
 */
@RestController
public class ExcelController {


    /**
     *
     * @param response  浏览器相应
     * @param request  获取前端传入的值 测试暂时未用
     * @return String
     */
    @RequestMapping("/exportExcel")
    @ResponseBody
    public  String exportExcel(HttpServletResponse response, HttpServletRequest request){
        try{
            response.setContentType("application/binary;charset=UTF-8");
            ServletOutputStream out=response.getOutputStream();
            //设置文件头：最后一个参数是设置下载文件名(这里我们叫：张三.pdf)
            String excelName="导出的数据表格";
            response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode(excelName+".xls", "UTF-8"));
            //用作传入的参数不正确时，返回给浏览器的提示信息
            if (StringUtils.isEmpty(excelName)){
                response.getWriter().print("传入参数不全!");
                return "传入参数不全";
            }
           //拼装测试数据
            LinkedMap<String,String> titleMap=new LinkedMap<>();
            titleMap.put("name","姓名");
            titleMap.put("age","年龄");
            titleMap.put("address","居住地址");
            titleMap.put("email","邮箱");
            titleMap.put("telephone","电话号码");
            List<User> list = new ArrayList<>();
            list.add(new User("张三","28","成都市","zhangsan@126.com","13312345678"));
            list.add(new User("李四","36","中江县","lisi@163.com","18998765432"));
            list.add(new User("王五","75","北京市","wangwu@126.com","15676543456"));
            list.add(new User("赵六","47","上海市","zhaoliu@126.com","1387656789"));
            ExcelUtils.exportExcel(titleMap, out,list);
            return "success";
        } catch(Exception e){
            return "导出信息失败";
        }
    }
}


