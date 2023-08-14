package com.heroku.java;


import java.io.File;
import java.util.Calendar;

public class PathHelper {
    public static final String CURRENT_DIR = ".";
    public static final String PROJECT_DIR_PICTURE = "/picture";
    public static final String PROJECT_DIR_SAVE_PICTURE_USER = "/picture/user";
    public static final String RESOURCE_DIR = "/src/main/resources/templates";
    public static final String TEMPLATE_DIR_SAVE_FILE = "/temp";
    public static final String TEMPLATE_DIR = "/target/classes/templates";
    public static final String TARGET_DIR = "/target";
    public static final String CLASSES_DIR = "/target/classes";
    public static final String TEMPLATE_FILE_EXCEL_EN = "Rikkeisoft_Skill_Sheet_Template_en.xlsx";
    public static final String TEMPLATE_FILE_EXCEL_JA = "Rikkeisoft_Skill_Sheet_Template_ja.xlsx";
    public static final String MEDIA_TYPE_APP_VND_EXCEL= "application/vnd.ms-excel";
    public static final String HEADER_ATTACHMENT_FILE = "attachment;filename=";
    public static final String EXTENSION = ".xlsx";

    public static  String getCurrentPath(){
        File currDir = new File(CURRENT_DIR);
        String path = currDir.getAbsolutePath();
        path = path.substring(0,path.length()-2);
        return path;
    }
    public static String getProjectDirPicture( String path){
        String paths = null;
        File current = new File(PROJECT_DIR_PICTURE);
        if (path.isEmpty()){
            // get default path
            if (current.exists()){
                if (current.isDirectory()){
                    paths = current.getAbsolutePath();
                }
            }

        }else {
            if (current.exists()){
                if (current.getName().equals(path)){
                    if (current.isDirectory()){
                        paths = current.getAbsolutePath();
                    }
                }else {
                    File[] files = current.listFiles();
                    if (files != null){
                        for (File f:files
                        ) {
                            if (f.isDirectory()){
                                paths = getFolder(f,path);
                                if (paths != null){
                                    return paths;
                                }
                            }
                        }
                    }
                }

            }
        }
        return paths;
    }


    private static String getFolder( File file, String path) {
        if (file.getName().equals(path)){
            return file.getAbsolutePath();
        }
        File[] files = file.listFiles();
        if (files != null){
            for (File f:files
            ) {
                if (f.isDirectory()){
                    return getFolder(f,path);
                }
            }
        }
        return createDir(path);
    }

    private static String createDir(String name){
        String path = getCurrentPath();
        File file = new File(path+"/"+name);
        if (!file.exists()) {
            if(file.mkdirs()){
                return file.getAbsolutePath();
            }
        }
        return  file.getAbsolutePath();
    }
    public static String createDirInPictureDir(String name){
        String path = getCurrentPath();
        File file = new File(path+"/"+PROJECT_DIR_PICTURE+"/"+name);
        if (!file.exists()){
            if(file.mkdirs()){
                return name;
            }
        }
        return name;
    }

    public static String createDirByDate(String name){

        Calendar cal = Calendar.getInstance();
        int year = cal.get(Calendar.YEAR);
        int month = cal.get(Calendar.MONTH)+1;
        int day = cal.get(Calendar.DATE);
        String date = Integer.toString(year)+"_"+Integer.toString(month)+"_"+Integer.toString(day);
        return createDirInPictureDir(date+"/"+name);


    }

}
