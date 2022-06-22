package com.lg.project.notes.common.exception;

/**
 * 异常回调类
 * @author Administrator
 * @param <T>
 */
public abstract class AbstractExceptionCallback<T> {

    protected abstract T doInTry() throws Exception;

    protected void doInCatch(Exception ex){
    }

    protected void doInFinally(){
    }

    protected boolean dropException(Exception ex){
        return false;
    }
}
