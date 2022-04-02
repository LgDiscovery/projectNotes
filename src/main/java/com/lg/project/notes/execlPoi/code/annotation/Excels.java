package com.lg.project.notes.execlPoi.code.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @ClassName Excels
 * @Description Excel注解集
 * @Author liuguang
 * @Date 2022/4/2 17:32
 * @Version 1.0
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Excels {

   public Excel[] value();
}
