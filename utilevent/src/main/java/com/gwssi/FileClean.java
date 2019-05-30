package com.gwssi;

import com.gwssi.util.PropertiesUtil;
import org.apache.log4j.Logger;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.Date;

/**
 * @Author LiuFei
 * @Date 2019/4/2 16:57
 * @description  用于指定文件夹中文件的清理
 */
@Component
public class FileClean {


    private static Logger log = Logger.getLogger(FileClean.class);


    @Scheduled(cron = "0 0 0 1/1 * ? ")
    public void deleteFile(){

        String fileUrl = new PropertiesUtil().getProperty("FILE_URL");

        //判断是否指定了多个url
        if(fileUrl.indexOf("|") == -1){
            log.info("文件清理开始,本次指定了1个Url...");
            String[] params = fileUrl.split("\\*");
            checkFile(params[0], Long.parseLong(params[1]));
        }else{
            //循环多个指定的url
            String[] fileUrls = fileUrl.split("\\|");
            log.info("文件清理开始,本次指定了"+fileUrls.length+"个Url...");
            for (int i=0;i<fileUrls.length;i++) {
                String[] params = fileUrls[i].split("\\*");
                checkFile(params[0], Long.parseLong(params[1]));
            }
        }

    }

    public void checkFile(String url, long day) {
        log.info("url:"+url+"开始检查超过"+ day +"天的文件...");
        File file = new File(url);
        File[] files = file.listFiles();

        for(int i=0;i<files.length;i++){
            file = files[i];
            if(file.isDirectory()){
                //如果是文件夹  递归
                checkFile(file.getPath(), day);
            }else{
                String fileName = file.getName();
                try {
                    //判断文件是否为catalina.out
                    if("catalina.out".equals(fileName)){
                        Runtime runtime = Runtime.getRuntime();
                        //获取tomcat根路径
                        String catalinaHome = System.getProperty("catalina.home");
                        String shFilePath = catalinaHome+"/shell/fileClean.sh";

                        File shDirPath = new File(catalinaHome+"/shell");
                        if(!shDirPath.exists()){
                            shDirPath.mkdir();
                        }
                        FileWriter fw = new FileWriter(new File(shFilePath));
                        //写入清空文件脚本
                        fw.write("truncate -s 0 "+file.getPath());
                        log.info("写入shell脚本:"+"truncate -s 0 "+file.getPath());
                        fw.close();
                        runtime.exec("chmod 777 "+shFilePath);
                        runtime.exec("/bin/sh "+shFilePath);
                        log.info("执行shell脚本:"+"/bin/sh "+shFilePath);
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }
                //判断文件最后修改时间是否超过指定天数
                long fileTime = file.lastModified()+day*24*60*60*1000;
                if(fileTime < new Date().getTime()){
                    log.info("文件:"+fileName+"删除!!!");
                    file.delete();
                }
            }
        }
    }

}
