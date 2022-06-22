package com.lg.project.notes.common.exception;

/**
 * 异常模板类
 * @author Administrator
 */
public final class ExceptionTemplate {

    public static void execute(AbstractExceptionCallback<?> callback) throws Exception {
        executeAndReturn(callback);
    }

    public static <T> T executeAndReturn(AbstractExceptionCallback<T> callback) throws Exception {
        try {
            return callback.doInTry();
        }catch (Exception e){
            if(callback.dropException(e)){
                return null;
            }else{
                callback.doInCatch(e);
                throw new Exception(e);
            }
        }finally {
            callback.doInFinally();
        }
    }
}
