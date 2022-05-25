package com.lg.project.notes;

import com.lg.project.notes.common.utils.LockUtils;

/**
 * @ClassName LockUtilsTest
 * @Description test1 方法测试了重入锁的效果。
 *               test2 测试了主线程获取锁之后一直未释放，持有锁超时之后被 thread1 获取到了。
 * @Author liuguang
 * @Date 2022/5/25 14:06
 * @Version 1.0
 */
public class LockUtilsTest {

    public void test1() throws Exception {
        String lock_key = "key1";
        for (int i = 0; i < 10; i++) {
            LockUtils.lock(lock_key, 10000L, 1000);
        }
        for (int i = 0; i < 9; i++) {
            LockUtils.unlock(lock_key);
        }
    }

    public void test2() throws Exception {
        String lock_key = "key2";
        LockUtils.lock(lock_key, 5000L, 1000);
        Thread thread1 = new Thread(() -> {
            try {
                try {
                    LockUtils.lock(lock_key, 5000L, 7000);
                } finally {
                    LockUtils.unlock(lock_key);
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        });
        thread1.setName("thread1");
        thread1.start();
        thread1.join();
    }

}
