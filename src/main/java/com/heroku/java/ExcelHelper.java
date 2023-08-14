package com.heroku.java;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;
@Component

public class ExcelHelper {
    public static byte[] excelExport(Candidate candidate, FileInputStream fis, String lang) {


        // use for project list table
//        Set<Project> projectList = new HashSet<>();


        try {
            //  Candidate
            Workbook workbook = new XSSFWorkbook(fis);
            // set up style, print, ...

            // Create a new font and alter it.
            Font font = workbook.createFont();
            Font font11Bold = workbook.createFont();
            font.setFontHeightInPoints((short) 11);
            font.setFontName("Times New Romance");
            font.setBold(false);
            font11Bold.setFontHeightInPoints((short) 11);
            font11Bold.setFontName("Times New Romance");
            font11Bold.setBold(true);

            CellStyle style = workbook.createCellStyle();
            CellStyle styleRight = workbook.createCellStyle();
            CellStyle styleCenter = workbook.createCellStyle();
            style.setWrapText(true);
            styleRight.setWrapText(true);
            styleCenter.setWrapText(true);

            // styleCenter.setFont(font);
            styleCenter.setAlignment(HorizontalAlignment.CENTER);
            styleCenter.setVerticalAlignment(VerticalAlignment.CENTER);
            styleCenter.setBorderRight(BorderStyle.THIN);
            styleCenter.setBorderLeft(BorderStyle.THIN);
            styleCenter.setBorderTop(BorderStyle.THIN);
            styleCenter.setBorderBottom(BorderStyle.THIN);
            styleCenter.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            styleCenter.setTopBorderColor(IndexedColors.BLACK.getIndex());
            styleCenter.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            styleCenter.setRightBorderColor(IndexedColors.BLACK.getIndex());


            style.setAlignment(HorizontalAlignment.LEFT);
            style.setVerticalAlignment(VerticalAlignment.CENTER);

            styleRight.setAlignment(HorizontalAlignment.RIGHT);
            styleRight.setVerticalAlignment(VerticalAlignment.CENTER);
            //set border
            setCellStyle(style);

            setCellStyle(styleRight);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setFont(font);

            Sheet sheet = workbook.getSheetAt(0);
//            if (candidate.getProjects()!=null){
//                projectList.addAll(candidate.getProjects()) ;
//                projectList = projectList.stream().sorted(Comparator.comparing(Project::getProjectNo)).collect(Collectors.toCollection(LinkedHashSet::new));
//            }

            // File currDir = new File(RankingType.CURRENT_DIR);
//            String path = PathHelper.getCurrentPath();
//            String fileLocation = path +PathHelper.RESOURCE_DIR + PathHelper.TEMPLATE_DIR_SAVE_FILE + File.separator
//                    + RankingType.deAccent(candidate.getName()) + "_" + lang ;
            //setup candidate
            Row currentRow = sheet.getRow(RankingType.NAME_CANDIDATE_ROW);
            Cell current = currentRow.getCell(RankingType.CANDIDATE_NAME_CELL);
            current.setCellValue(candidate.getName());
            current = currentRow.getCell(RankingType.CANDIDATE_NAME_KANA_CELL);
            current.setCellValue(candidate.getNameKana());
//            setImageToCell(sheet, workbook, candidate.getImage(), 8, RankingType.NAME_CANDIDATE_ROW);
            current = currentRow.getCell(RankingType.CANDIDATE_AGE_CELL);
            if (lang.equals(RankingType.FILE_TYPE_EN)) {
                current.setCellValue(candidate.isGender() ? RankingType.MALE_EN : RankingType.FEMALE_EN);
            } else {
                current.setCellValue(candidate.isGender() ? RankingType.MALE_JA : RankingType.FEMALE_JA);
            }
            //
            currentRow = sheet.getRow(RankingType.ALMA_MASTER_ROW);
            current = currentRow.getCell(RankingType.CANDIDATE_ALMA_CELL);
            current.setCellValue(candidate.getGraduateSchool());
            current = currentRow.getCell(RankingType.CANDIDATE_AGE_CELL);
            //do age
            Calendar cal = Calendar.getInstance();
            if (candidate.getBirthDay()== null){
                current.setCellValue(53);
            }else {
                cal.setTime(candidate.getBirthDay());
                Calendar birthday = Calendar.getInstance();
                birthday.setTime(new Date());
                int yearCurrent = birthday.get(Calendar.YEAR);
                int yearBirth = cal.get(Calendar.YEAR);
                current.setCellValue(yearCurrent-yearBirth);

            }


            currentRow = sheet.getRow(RankingType.COMPATIBLE_FIELD_ROW);
            current = currentRow.getCell(RankingType.CANDIDATE_COMPATIBLE_CELL);
            current.setCellValue(candidate.getCompatibleField());
            current = currentRow.getCell(RankingType.CANDIDATE_ENGLISH_CELL);

//            if (candidate.getEnglishLevel()== null){
//                current.setCellValue("");
//            } else {
//                current.setCellValue(candidate.getEnglishLevel().getDescription());
//
//            }

            currentRow = sheet.getRow(RankingType.WORK_EXPERIENCE_ROW);
            current = currentRow.getCell(RankingType.CANDIDATE_EXPERIENCE_CELL);
            if (lang.equals(RankingType.FILE_TYPE_EN)) {
                current.setCellValue(candidate.getExperienceYears() + RankingType.YEAR_EN);
            } else {
                current.setCellValue(candidate.getExperienceYears() + RankingType.YEAR_JA);
            }

            current = currentRow.getCell(RankingType.CANDIDATE_EXP_JAPAN_CELL);
            current.setCellValue(candidate.getExperienceYearsInJapan());
            current = currentRow.getCell(RankingType.CANDIDATE_JAPANESE_CELL);

//            if (candidate.getJapaneseLevel()== null){
//                current.setCellValue("");
//            }else {
//                current.setCellValue(candidate.getJapaneseLevel().getDescription());
//            }


            currentRow = sheet.getRow(RankingType.PERSONAL_ROW);
            current = currentRow.getCell(RankingType.CANDIDATE_PERSONAL_CELL);
            current.setCellValue(candidate.getSelfPR());
            int firstRow = RankingType.FIRST_PROJECT_ROW;
            int lastRow = RankingType.LAST_PROJECT_ROW;
//            if (projectList.size()>0) {
//                int rowCount = 0;
//                for (Project p : projectList) {
//                    CellRangeAddress rangeAddress = new CellRangeAddress(firstRow, lastRow, 0, 0);
//
//                    sheet.addMergedRegion(rangeAddress);
//                    doMerge(rangeAddress, sheet);
//                    sheet.addMergedRegion(new CellRangeAddress(firstRow, firstRow, 1, 2));
//                    doMerge(new CellRangeAddress(firstRow, firstRow, 1, 2), sheet);
//                    sheet.addMergedRegion(new CellRangeAddress(firstRow + 1, lastRow, 1, 2));
//                    doMerge(new CellRangeAddress(firstRow + 1, lastRow, 1, 2), sheet);
//                    sheet.addMergedRegion(new CellRangeAddress(firstRow + 1, lastRow, 3, 3));
//                    doMerge(new CellRangeAddress(firstRow + 1, lastRow, 3, 3), sheet);
//                    sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, 4, 4));
//                    doMerge(new CellRangeAddress(firstRow, lastRow, 4, 4), sheet);
//                    sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, 5, 5));
//                    doMerge(new CellRangeAddress(firstRow, lastRow, 5, 5), sheet);
//                    sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, 6, 6));
//                    doMerge(new CellRangeAddress(firstRow, lastRow, 6, 6), sheet);
//                    sheet.addMergedRegion(new CellRangeAddress(firstRow, firstRow + 1, 7, 7));
//                    doMerge(new CellRangeAddress(firstRow, firstRow + 1, 7, 7), sheet);
//                    sheet.addMergedRegion(new CellRangeAddress(firstRow + 2, lastRow, 7, 7));
//                    doMerge(new CellRangeAddress(firstRow + 2, lastRow, 7, 7), sheet);
//                    sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, 8, 8));
//                    doMerge(new CellRangeAddress(firstRow, lastRow, 8, 8), sheet);
//
//                    // get project role name
//                    String projectRoleName = projectRoles.stream().filter(item -> item.getId() == p.getProjectRole()).findFirst().get().getDescription();
//
//                    //set data to cells
//                    rowCount = firstRow;
//                    for (int i = 1; i <= 4; i++) {
//                        Row row = sheet.getRow(rowCount);
//                        row.setHeightInPoints((short) 25);
//                        Cell cell;
//                        for (int j = 0; j < 8; j++) {
//                            if (rowCount % 2 == 1) {
//                                switch (j) {
//                                    case RankingType.CANDIDATE_PROJECT_NO_CELL:
//                                        cell = row.createCell(j, CellType.NUMERIC);
//                                        cell.setCellStyle(style);
//                                        cell.setCellValue(p.getProjectNo());
//                                        break;
//                                    case RankingType.CANDIDATE_PROJECT_NAME_CELL:
//                                        cell = row.createCell(j, CellType.STRING);
//                                        cell.setCellStyle(style);
//                                        cell.setCellValue(p.getProjectName());
//                                        break;
//                                    case RankingType.CANDIDATE_PROJECT_ROLE_CELL:
//                                        cell = row.createCell(j, CellType.STRING);
//                                        cell.setCellStyle(style);
//                                        cell.setCellValue(projectRoleName);
//                                        break;
//                                    case RankingType.CANDIDATE_PROJECT_PROGRAMING_CELL:
//                                        //set programming languages
//                                        cell = row.createCell(j, CellType.STRING);
//                                        cell.setCellStyle(style);
//                                        cell.setCellValue(getStringFromListId(skillList,p.getProgrammingLanguage()));
//                                        break;
//                                    case RankingType.CANDIDATE_PROJECT_ENVIRONMENT_CELL: //set environment
//                                        //project.setEnvironment();
//                                        cell = row.createCell(j, CellType.STRING);
//                                        cell.setCellStyle(style);
//                                        cell.setCellValue(getStringFromListId(skillList,p.getEnvironment()));
//                                        break;
//                                    case RankingType.CANDIDATE_PROJECT_PHASES_CELL: // set phases
//                                        cell = row.createCell(j, CellType.STRING);
//                                        cell.setCellStyle(style);
//                                        cell.setCellValue(getStringFromListId(skillList,p.getProjectPhases()));
//                                        break;
//                                    case RankingType.CANDIDATE_PROJECT_START_DATE_CELL: //set start date
//                                        cell = row.createCell(j, CellType.STRING);
//                                        cell.setCellStyle(styleCenter);
//                                        cell.setCellValue(getStrFromDate(p.getStartDate(), lang));
//                                        //set period
//                                        cell = row.createCell(j + 1, CellType.STRING);
//                                        cell.setCellValue(getPeriod(p.getStartDate(), p.getEndDate()));
//                                        cell.setCellStyle(styleCenter);
//                                        break;
//
//                                }
//                            } else {
//
//                                switch (j) {
//                                    case RankingType.CANDIDATE_PROJECT_DESCRIPTION_CELL:
//                                        cell = row.createCell(j, CellType.STRING);
//                                        cell.setCellStyle(style);
//                                        cell.setCellValue(p.getProjectDescription());
//                                        break;
//
//                                    case RankingType.CANDIDATE_PROJECT_MEMBERS_CELL:
//                                        cell = row.createCell(j, CellType.STRING);
//                                        cell.setCellStyle(style);
//                                        cell.setCellValue(RankingType.getTeam(p.getProjectMembers(), p.getProjectScale(), lang));
//                                        break;
//
//                                    case RankingType.CANDIDATE_PROJECT_END_DATE_CELL:
//                                        if (p.getEndDate() != null) {
//                                            cell = row.createCell(j, CellType.STRING);
//                                            cell.setCellStyle(styleCenter);
//                                            cell.setCellValue(getStrFromDate(p.getEndDate(), lang));
//                                        }
//
//                                        break;
//
//                                }
//                            }
//
//                        }
//                        rowCount++;
//                    }
//
//                    firstRow = firstRow + 4;
//                    lastRow = lastRow + 4;
//                }
//            }
            // merge reference
            firstRow = firstRow +1;
            lastRow = lastRow +2;
            sheet.addMergedRegion(new CellRangeAddress(firstRow,firstRow+2,0,1));
            doMerge(new CellRangeAddress(firstRow,firstRow+2,0,1),sheet);
            sheet.addMergedRegion(new CellRangeAddress(firstRow,firstRow+2,2,8));
            doMerge(new CellRangeAddress(firstRow,firstRow+2,2,8),sheet);
            Row row = sheet.getRow(firstRow);
            if (lang.equals(RankingType.FILE_TYPE_EN)){
                Cell c=  row.createCell(0);
                c.setCellStyle(style);
                c.setCellValue(RankingType.REFERENCE_EN);
                setCellStyleBorderLeft(c, workbook.createCellStyle(), workbook.createFont());
            }else {
                Cell c=  row.createCell(0);
                c.setCellStyle(style);
                c.setCellValue(RankingType.REFERENCE_JA);
                setCellStyleBorderLeft(c, workbook.createCellStyle(), workbook.createFont());
            }
            Cell c=  row.createCell(2);
            c.setCellValue(candidate.getReference());
            c.setCellStyle(style);

            // set up ranking type
//            Set<CandidateSkill> skillSet = new HashSet<>();
//            skillSet.addAll(candidate.getSkills().stream().filter(key->key.getSkill().getCategory().getId() == ConstantHelper.OS).collect(Collectors.toList()));
//            int firstRowRank = 9; // from  9th row
//            int firstCellRank = 10; // from 10th cell
//            if (skillSet.size()>0){
//                List<Skill> skillAll =  getSkillById(skillList,4);
//                Row row1 = sheet.getRow(firstRowRank);
//                Cell cell = row1.createCell(firstCellRank,CellType.STRING);
//                sheet.addMergedRegion(new CellRangeAddress(firstRowRank,firstRowRank,10,16));
//                doMerge(new CellRangeAddress(firstRowRank,firstRowRank,10,16),sheet);
//                cell.setCellValue(RankingType.OS_SKILL_JA);
//                setCellStyleBorderLeft(cell, workbook.createCellStyle(),workbook.createFont());
//                firstRowRank++;
//                for (CandidateSkill os :skillSet
//                ) {
//
//                    row1 = sheet.getRow(firstRowRank);
//                    row1.setHeightInPoints((short)25);
//
//                    setSkillSheet(lang, style, styleRight, styleCenter, skillAll, row1, os);
//                    firstRowRank++;
//
//                }
//
//            }
            // database skill
//            skillSet.clear();
//            skillSet.addAll(candidate.getSkills().stream().filter(key->key.getSkill().getCategory().getId() == ConstantHelper.DATABASE).collect(Collectors.toList()));
//
//            if (skillSet.size()>0) {
//                List<Skill> skillAll =  getSkillById(skillList,3);;
//                Row row1 = sheet.getRow(firstRowRank);
//                Cell cell = row1.createCell(firstCellRank,CellType.STRING);
//                sheet.addMergedRegion(new CellRangeAddress(firstRowRank,firstRowRank,10,16));
//                doMerge(new CellRangeAddress(firstRowRank,firstRowRank,10,16),sheet);
//                cell.setCellValue(RankingType.DATABASE_SKILL_JA);
//                setCellStyleBorderLeft(cell, workbook.createCellStyle(),workbook.createFont());
//                firstRowRank++;
//                for (CandidateSkill os :skillSet
//                ) {
//                    row1 = sheet.getRow(firstRowRank);
//                    row1.setHeightInPoints((short)25);
//                    firstRowRank++;
//                    setSkillSheet(lang, style, styleRight, styleCenter, skillAll, row1, os);
//
//                }
//
//            }
            // framework skill
//            skillSet.clear();
//            skillSet.addAll(candidate.getSkills().stream().filter(key->key.getSkill().getCategory().getId() == ConstantHelper.IDE).collect(Collectors.toList()));
//            if (skillSet.size()>0){
//                {
//                    List<Skill> skillAll =  getSkillById(skillList,2);;
//                    Row row1 = sheet.getRow(firstRowRank);
//                    Cell cell = row1.createCell(firstCellRank,CellType.STRING);
//                    sheet.addMergedRegion(new CellRangeAddress(firstRowRank,firstRowRank,10,16));
//                    doMerge(new CellRangeAddress(firstRowRank,firstRowRank,10,16),sheet);
//                    cell.setCellValue(RankingType.FRAMEWORK_SKILL_JA);
//                    setCellStyleBorderLeft(cell, workbook.createCellStyle(),workbook.createFont());
//                    firstRowRank++;
//                    for (CandidateSkill os :skillSet
//                    ) {
//                        row1 = sheet.getRow(firstRowRank);
//                        row1.setHeightInPoints((short)25);
//                        firstRowRank++;
//                        setSkillSheet(lang, style, styleRight, styleCenter, skillAll, row1, os);
//
//                    }
//
//                }
//            }
            // programming languages skill
//            skillSet.clear();
//            skillSet.addAll(candidate.getSkills().stream().filter(key->key.getSkill().getCategory().getId() == ConstantHelper.PROGRAMMING_LANGUAGE).collect(Collectors.toList()));
//
//            if (skillSet.size()>0) {
//                {
//                    List<Skill> skillAll =  getSkillById(skillList,1);;
//                    Row row1 = sheet.getRow(firstRowRank);
//                    Cell cell = row1.createCell(firstCellRank,CellType.STRING);
//                    sheet.addMergedRegion(new CellRangeAddress(firstRowRank,firstRowRank,10,16));
//                    doMerge(new CellRangeAddress(firstRowRank,firstRowRank,10,16),sheet);
//                    cell.setCellValue(RankingType.PROGRAMMING_LANGUAGE_SKILL_JA);
//                    setCellStyleBorderLeft(cell, workbook.createCellStyle(),workbook.createFont());
//
//                    firstRowRank++;
//                    for (CandidateSkill os : skillSet
//                    ) {
//                        row1 = sheet.getRow(firstRowRank);
//                        row1.setHeightInPoints((short)25);
//                        firstRowRank++;
//                        setSkillSheet(lang, style, styleRight, styleCenter, skillAll, row1, os);
//
//                    }
//
//                }
//            }
            // english  skill
//            skillSet.clear();
//            skillSet.addAll(candidate.getSkills().stream().filter(key->key.getSkill().getCategory().getId() == ConstantHelper.ENGLISH_SKILL).collect(Collectors.toList()));
//            if (skillSet.size()>0) {
//                {
//                    List<Skill> skillAll =  getSkillById(skillList,5);;
//                    Row row1 = sheet.getRow(firstRowRank);
//                    Cell cell = row1.createCell(firstCellRank,CellType.STRING);
//                    sheet.addMergedRegion(new CellRangeAddress(firstRowRank,firstRowRank,10,16));
//                    doMerge(new CellRangeAddress(firstRowRank,firstRowRank,10,16),sheet);
//                    cell.setCellValue(RankingType.ENGLISH_SKILL_JA);
//                    setCellStyleBorderLeft(cell, workbook.createCellStyle(), workbook.createFont());
//                    firstRowRank++;
//                    for (CandidateSkill os :skillSet
//                    ) {
//                        row1 = sheet.getRow(firstRowRank);
//                        row1.setHeightInPoints((short)25);
//                        firstRowRank++;
//                        for (int i =10; i<=16; i++){
//                            cell = row1.createCell(i,CellType.STRING);
//                            cell.setCellStyle(styleCenter);
//                            switch (i){
//                                case RankingType.RANKING_NAME_CELL:
//                                    cell = row1.createCell(i,CellType.STRING);
//                                    cell.setCellStyle(style);
//                                    //Optional<Skill> name = skillAll.stream().filter(it->it.getId()== os.getId()).findFirst();
//                                    // if (name.isPresent()){
//                                    cell.setCellValue(os.getSkill().getDescription());
//                                    //  }
//                                    break;
//
//                                case RankingType.YEAR_EXPERIENCE_CELL:
//                                    cell =  row1.createCell(i,CellType.STRING);
//                                    cell.setCellStyle(styleRight);
//
//                                    if (lang.equals(RankingType.FILE_TYPE_EN)){
//                                        cell.setCellValue((os.getYears()+ RankingType.YEAR_EN+" - "+ os.getMonths() +RankingType.MONTH_EN_));
//
//                                    }else {
//                                        cell.setCellValue((os.getYears()+ RankingType.YEAR_JA_+" - "+ os.getMonths() +RankingType.MONTH_JA_));
//
//                                    }
//
//                                    break;
//                                default:
//
//                                    setRowCellRanking(row1, os.getLevel(),styleCenter);
//                                    break;
//
//                            }
//                        }
//
//                    }
//
//                }
//            }




            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            workbook.close();
//            try (OutputStream outputStream = new FileOutputStream(fileLocation)){
//                workbook.write(outputStream);
//                workbook.close();
//                outputStream.close();
//            }

            fis.close();
//            workbook.close();


            return outputStream.toByteArray();
        } catch (IOException e) {
            return null;

        }


    }
    private static void setCellStyle(CellStyle style) {
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());

        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());

        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());

        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());


        style.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
        style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
    }
    public static void doMerge(CellRangeAddress rang, Sheet sheet) {

        RegionUtil.setBorderBottom(BorderStyle.THIN, rang, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, rang, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, rang, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, rang, sheet);

        RegionUtil.setLeftBorderColor(IndexedColors.BLACK.getIndex(), rang, sheet);
        RegionUtil.setBottomBorderColor(IndexedColors.BLACK.getIndex(), rang, sheet);
        RegionUtil.setRightBorderColor(IndexedColors.BLACK.getIndex(), rang, sheet);
        RegionUtil.setTopBorderColor(IndexedColors.BLACK.getIndex(), rang, sheet);

    }

    private static void setCellStyleBorderLeft(Cell cell, CellStyle style, Font font ){
        font.setBold(true);
        style.setWrapText(true);
        style.setFont(font);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(style);
    }

}
