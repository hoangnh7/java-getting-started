package com.heroku.java;


import java.text.Normalizer;
import java.util.regex.Pattern;

public final class RankingType {
    /*public static final int RANKING_ROW = 9;
    public static final int FROM_RANKING_CELL = 10;*/
    public static final int RANKING_NAME_CELL = 10;
    public static final String LANGUAGE_PARAM ="lang";
    public static final int YEAR_EXPERIENCE_CELL = 16;
    public static final int LEVEL_1 = 11;
    public static final int LEVEL_2 = 12;
    public static final int LEVEL_3 = 13;
    public static final int LEVEL_4 = 14;
    public static final int LEVEL_5 = 15;

    static final int NAME_CANDIDATE_ROW = 3;
    static final int ALMA_MASTER_ROW = 4;

    static final int WORK_EXPERIENCE_ROW = 6;
    static final int COMPATIBLE_FIELD_ROW = 5;
    static final int PERSONAL_ROW = 8;

    static final int FIRST_PROJECT_ROW = 13;
    static final int LAST_PROJECT_ROW = 16;
    static final int CANDIDATE_NAME_CELL = 2;

    static final int CANDIDATE_COMPATIBLE_CELL = 2;
    static final int CANDIDATE_NAME_KANA_CELL = 5;
    static final int CANDIDATE_GENDER_CELL = 7;
    static final int CANDIDATE_ALMA_CELL = 2;
    static final int CANDIDATE_AGE_CELL = 7;

    static final int CANDIDATE_ENGLISH_CELL = 7;

    static final int CANDIDATE_EXPERIENCE_CELL = 2;

    static final int CANDIDATE_EXP_JAPAN_CELL = 5;
    static final int CANDIDATE_JAPANESE_CELL = 7;

    static final int CANDIDATE_PERSONAL_CELL = 2;

    static final int CANDIDATE_PROJECT_NO_CELL = 0;
    static final int CANDIDATE_PROJECT_NAME_CELL = 1;
    static final int CANDIDATE_PROJECT_ROLE_CELL = 3;

    static final int CANDIDATE_PROJECT_PROGRAMING_CELL = 4;
    static final int CANDIDATE_PROJECT_ENVIRONMENT_CELL = 5;
    static final int CANDIDATE_PROJECT_PHASES_CELL = 6;
    static final int CANDIDATE_PROJECT_START_DATE_CELL = 7;
    static final int CANDIDATE_PROJECT_DESCRIPTION_CELL = 1;
    static final int CANDIDATE_PROJECT_MEMBERS_CELL = 3;
    static final int CANDIDATE_PROJECT_END_DATE_CELL = 7;

    static final String PROGRAMMING_LANGUAGE_SKILL_EN = "Languages";
    static final String PROGRAMMING_LANGUAGE_SKILL_JA = "言語";
    static final String REFERENCE_EN = "Reference";
    static final String REFERENCE_JA = "参照";

    static final String OS_SKILL_EN = "OS";
    static final String OS_SKILL_JA = "OS";
    static final String DATABASE_SKILL_EN = "Database";
    static final String DATABASE_SKILL_JA = "データベース";
    static final String FRAMEWORK_SKILL_EN = "Framework";
    static final String FRAMEWORK_SKILL_JA = "フレームワーク";

    static final String ENGLISH_SKILL_EN = "English";
    static final String ENGLISH_SKILL_JA = "英語";
    public static final String FILE_NAME_PARAM = "file";
    static final int LANGUAGES_STATUS = 1;
    static final int LEVEL_SKILL_1 = 1;
    static final int LEVEL_SKILL_2 = 2;
    static final int LEVEL_SKILL_3 = 3;
    static final int LEVEL_SKILL_4 = 4;
    static final int LEVEL_SKILL_5 = 5;
    static final int OS_STATUS = 4;
    static final int FRAMEWORK_STATUS = 2;

    static final int DATABASE_STATUS = 3;
    static  final int ENGLISH_STATUS = 5;
    static final String FILE_TEMPLATE_EXCEL_PATH = "target\\classes\\templates\\";
    static final String CHECK_POINT_JA = " ● ";

    static final String PRE_FILE_EXCEL_TEMPLATE = "Rikkeisoft_Skill_Sheet_";
    static final String EXTENSION_FILE = ".xlsx";
    static final String FILE_TYPE_EN = "en";

    static final String FILE_TYPE_JA = "ja";
    static final String MALE_JA ="男";
    static final String MALE_EN ="Male";

    static final String FEMALE_JA ="女";
    static final String FEMALE_EN ="Female";

    static final  String YEAR_JA = "年";
    static final  String YEAR_EN = "　Years";

    static final String MONTH_JA_ =" ヶ月";
    static final String MONTH_JA ="月";
    static final  String MONTH_EN_ = " Months";
    static final  String YEAR_JA_ = "　年 ";
    static final String CURRENT_DIR = ".";
    private RankingType(){}
    public static String deAccent( String str) {
        if (str.isEmpty()) return "";
        String nfdNormalizedString = Normalizer.normalize(str, Normalizer.Form.NFD);
        Pattern pattern = Pattern.compile("\\p{InCombiningDiacriticalMarks}+");
        return pattern.matcher(nfdNormalizedString).replaceAll("").replaceAll(" ","_");
    }


    public static String getTeam(int teamSize, int members,  String lang){
        String str ="";
        if (lang.equals(RankingType.FILE_TYPE_EN)){
            str = "Team: " +teamSize+ " members\nTotal: "+String.format("%.2f",(double)members)+" MM";
        }else {
            str = "テーム："+teamSize + " 名　\n全体："+String.format("%.2f",(double)members) +" MM";
        }
        return str;
    }
    public static int getTeamSize( String str,  String lang){
        int size = 0;
        String s="";
        if (str.isEmpty())return size;
        if (lang.equals(RankingType.FILE_TYPE_EN)){
            s = str.substring(0,str.indexOf("member"));
        }else {
            s = str.substring(0,str.indexOf("名"));
        }
        size = getIntFromString(s);
        return size;
    }
    public static int getTotal( String str,  String lang){
        int size = 0;
        String s="";
        if (str.isEmpty())return size;
        try {
            if (lang.equals(RankingType.FILE_TYPE_EN)){
                s = str.substring(str.indexOf("Total")+"Total:".length(),str.length());
                if (s.length()>0){
                    s= s.substring(0,s.indexOf("."));
                }

            }
            else {
                s = str.substring(str.indexOf("全体")+"全体:".length(),str.length());
                if (s.length()>0){
                    s= s.substring(0,s.indexOf("."));
                }

            }
            size = getIntFromString(s);
        }catch (Exception e){
            return 0;
        }

        return size;
    }

    static int getIntFromString( String s) {
        char[] chars =s.toCharArray();
        StringBuilder sb = new StringBuilder();
        try {
            for (char c:chars
            ) {
                if (Character.isDigit(c)){
                    sb.append(c);
                }
            }


        }catch (Exception e){}

        return Integer.parseInt(sb.toString().isEmpty()?"0":sb.toString());
    }
}

