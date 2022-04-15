package com.lg.project.notes.common.utils;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileItemFactory;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.commons.CommonsMultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * @ClassName MultipartFileUtils
 * @Description 获取文件上传类型
 * @Author liuguang
 * @Date 2022/4/11 18:29
 * @Version 1.0
 */
public class MultipartFileUtils {

    public static void main(String[] args) {

        File file = new File("C:\\Users\\cocboo\\Desktop\\test.xlsx");

        System.out.println(file.getName().substring(0,file.getName().lastIndexOf(".")));

        FileItem fileItem = getMultipartFile(file, "test");

        MultipartFile multipartFile = new CommonsMultipartFile(fileItem);

    }

    // 调用
    private static FileItem getMultipartFile(File file, String fieldName) {
        FileItemFactory factory = new DiskFileItemFactory(16, null);
        FileItem item = factory.createItem(fieldName, "text/plain", true, file.getName());
        int bytesRead = 0;
        int len = 8192;
        byte[] buffer = new byte[len];
        try {
            FileInputStream fis = new FileInputStream(file);
            OutputStream os = item.getOutputStream();
            while ((bytesRead = fis.read(buffer, 0, len)) != -1) {
                os.write(buffer, 0, bytesRead);
            }
            os.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return item;
    }

}
