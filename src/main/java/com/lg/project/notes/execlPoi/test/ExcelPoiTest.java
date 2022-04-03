package com.lg.project.notes.execlPoi.test;

import com.lg.project.notes.common.model.AjaxResult;
import com.lg.project.notes.execlPoi.code.entity.SysDept;
import com.lg.project.notes.execlPoi.code.entity.SysUser;
import com.lg.project.notes.execlPoi.code.poi.ExcelUtil;

import java.util.ArrayList;
import java.util.List;

/**
 * @ClassName ExcelPoiTest
 * @Description 用来测试execlpoi功能的测试模块
 * @Author liuguang
 * @Date 2022/4/2 16:46
 * @Version 1.0
 */
public class ExcelPoiTest {

    /** 
     * @description:
     * @param: args
     * @return: void
     * @author liuguang
     * @date: 2022/4/2 17:08
     */ 
    public static void main(String[] args) {

        ExcelUtil<SysUser> excelUtil = new ExcelUtil(SysUser.class);
        AjaxResult ajaxResult =excelUtil.exportExcel(initData(),"我的测试数据","这是一个数据标题");
        System.out.println(ajaxResult);

    }

    public static List<SysUser> initData(){
        List<SysUser> list  = new ArrayList<>();
        SysUser sysUser = new SysUser();
        sysUser.setUserId(1l);
        sysUser.setUserName("一号任务");
        sysUser.setNickName("小花");
        sysUser.setEmail("123456");
        sysUser.setPhonenumber("157792805455");
        sysUser.setSex("0");

        SysDept sysDept = new SysDept();
        sysDept.setDeptName("一号部门");
        sysDept.setLeader("总理");
        sysUser.setDept(sysDept);
        list.add(sysUser);
        return list;
    }
}
